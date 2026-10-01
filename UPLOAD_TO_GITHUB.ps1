param(
    [switch]$PrepareOnly,
    [string]$RepositoryUrl = 'https://github.com/murry1239/AT_GENERATOR.git'
)
$ErrorActionPreference = 'Stop'

function Invoke-Git {
    param([string[]]$Arguments)
    & git @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Git failed: $($Arguments -join ' ')" }
}

try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Install Git for Windows first.' }
    # Explicit source-only list. Never git add the surrounding workspace.
    $files = @(
        '.gitignore', 'analyzer.py', 'app.py', 'pair_analysis.py', 'pair_runner.py', 'word_capture.py',
        'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md',
        'INSTALL.bat', 'RUN.bat', 'DIAGNOSE.bat', 'TEST.bat',
        'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1',
        'tests/test_pair_analysis.py', 'tests/test_alpha04.py',
        'tools/validate_sample.py', 'tools/build_release.py'
    )
    foreach ($name in $files) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name) -PathType Leaf)) {
            throw "Missing release source: $name"
        }
    }
    if (-not $PrepareOnly) {
        Write-Host "Default repository: $RepositoryUrl"
        $answer = Read-Host 'Press Enter to use it, or enter another GitHub repository HTTPS URL'
        if ($answer) { $RepositoryUrl = $answer.Trim() }
    }
    if ($RepositoryUrl -notmatch '^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
        throw 'Only a GitHub HTTPS repository URL is supported. Do not put a token in the URL.'
    }
    $repo = Join-Path $PSScriptRoot '.github-upload'
    $branch = 'codex/alpha-0.4'
    if (-not (Test-Path -LiteralPath $repo)) {
        $bundle = Join-Path $PSScriptRoot 'AT_GENERATOR_history.bundle'
        if (-not (Test-Path -LiteralPath $bundle)) { throw 'History bundle missing. Use the full distribution ZIP.' }
        Invoke-Git @('clone', '-b', 'main', $bundle, $repo)
        Invoke-Git @('-C', $repo, 'switch', '-c', $branch)
    }
    $actualRoot = (& git -C $repo rev-parse --show-toplevel)
    if ($LASTEXITCODE -ne 0 -or [IO.Path]::GetFullPath($actualRoot) -ne [IO.Path]::GetFullPath($repo)) {
        throw 'The upload directory is not an independent Git repository.'
    }
    $actualBranch = (& git -C $repo branch --show-current)
    if ($LASTEXITCODE -ne 0 -or $actualBranch -ne $branch) { throw "Expected branch: $branch" }
    Invoke-Git @('-C', $repo, 'merge-base', '--is-ancestor', 'c463368', 'HEAD')
    & git -C $repo diff --quiet
    if ($LASTEXITCODE -ne 0) { throw 'The upload copy has unstaged edits. Preserve/review them before rerunning.' }
    $historyPaths = @(& git -C $repo log '--format=' '--name-only' 'c463368..HEAD')
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect local upload history.' }
    foreach ($name in $historyPaths) {
        if ($name -and $name -notin $files) { throw "Unexpected file in local history: $name. Review before upload." }
    }
    # Refuse unexpected staged paths left by a manual edit in the upload copy.
    $staged = @(& git -C $repo diff --cached --name-only)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged paths.' }
    foreach ($name in $staged) {
        if ($name -and $name -notin $files) { throw "Unexpected staged file: $name. Review it manually." }
    }
    foreach ($name in $files) {
        $destination = Join-Path $repo $name
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $destination
    }
    Invoke-Git (@('-C', $repo, 'add', '--') + $files)
    Invoke-Git @('-C', $repo, 'diff', '--cached', '--stat')
    if ($PrepareOnly) {
        Write-Host 'Prepared source-only staging. No commit, network access, or push was performed.'
        exit 0
    }
    Write-Host "Repository: $RepositoryUrl"
    Write-Host "Branch: $branch (main is not overwritten)"
    Write-Host 'Only listed source, tests, and documentation will be committed. Review the list above.'
    if ((Read-Host 'Type UPLOAD to commit and push') -cne 'UPLOAD') {
        Write-Host 'Cancelled. Local source staging remains available.'
        exit 0
    }
    $userName = (& git -C $repo config user.name)
    $userEmail = (& git -C $repo config user.email)
    if (-not $userName) {
        $userName = Read-Host 'Commit author name (saved only in this upload repository)'
        if (-not $userName.Trim()) { throw 'Author name is required. Nothing was pushed.' }
        Invoke-Git @('-C', $repo, 'config', 'user.name', $userName.Trim())
    }
    if (-not $userEmail) {
        $userEmail = Read-Host 'Commit email (visible in Git history; use your GitHub noreply email for privacy)'
        if (-not $userEmail.Trim()) { throw 'Author email is required. Nothing was pushed.' }
        Invoke-Git @('-C', $repo, 'config', 'user.email', $userEmail.Trim())
    }
    Invoke-Git @('-C', $repo, 'remote', 'set-url', 'origin', $RepositoryUrl)
    & git -C $repo diff --cached --quiet
    if ($LASTEXITCODE -eq 1) {
        Invoke-Git @('-C', $repo, 'commit', '-m', 'feat: Alpha 0.4 content-only diff and bounded native Word image capture')
    } elseif ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged changes.' }
    Invoke-Git @('-C', $repo, 'push', '-u', 'origin', $branch)
    Write-Host 'Upload completed. Alpha 0.2 and 0.3 history is retained; no force push was used.'
} catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
