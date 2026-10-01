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
    $localFiles = @(
        '.gitignore', 'app.py', 'generator.py', 'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md',
        'INSTALL.bat', 'RUN.bat', 'DIAGNOSE.bat', 'TEST.bat', 'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1',
        'tests/test_generator.py', 'tools/build_release.py'
    )
    $files = @($localFiles | ForEach-Object { 'BD_GENERATOR/' + $_ })
    foreach ($name in $localFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name) -PathType Leaf)) { throw "Missing source: $name" }
    }
    if (-not $PrepareOnly) {
        Write-Host "Default repository: $RepositoryUrl"
        $answer = Read-Host 'Press Enter to use it, or enter another GitHub HTTPS repository URL'
        if ($answer) { $RepositoryUrl = $answer.Trim() }
    }
    if ($RepositoryUrl -notmatch '^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
        throw 'Only a GitHub HTTPS repository URL is supported. Do not put a token in the URL.'
    }
    $repo = Join-Path $PSScriptRoot '.github-upload'
    $branch = 'codex/byeondaesaenggi-alpha-0.9'
    if (-not (Test-Path -LiteralPath $repo)) {
        $bundle = Join-Path $PSScriptRoot 'AT_GENERATOR_history.bundle'
        if (-not (Test-Path -LiteralPath $bundle)) { throw 'History bundle missing. Use the full distribution ZIP.' }
        Invoke-Git @('clone', '-b', 'main', $bundle, $repo)
        Invoke-Git @('-C', $repo, 'switch', '-c', $branch)
    }
    $actualRoot = (& git -C $repo rev-parse --show-toplevel)
    if ($LASTEXITCODE -ne 0 -or [IO.Path]::GetFullPath($actualRoot) -ne [IO.Path]::GetFullPath($repo)) {
        throw 'Upload directory must be an independent Git repository.'
    }
    $actualBranch = (& git -C $repo branch --show-current)
    if ($LASTEXITCODE -ne 0 -or $actualBranch -ne $branch) { throw "Expected branch: $branch" }
    Invoke-Git @('-C', $repo, 'merge-base', '--is-ancestor', 'c463368', 'HEAD')
    & git -C $repo diff --quiet
    if ($LASTEXITCODE -ne 0) { throw 'Upload copy has unstaged edits. Preserve/review them before rerunning.' }
    $historyPaths = @(& git -C $repo log '--format=' '--name-only' 'c463368..HEAD')
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect upload history.' }
    $staged = @(& git -C $repo diff --cached --name-only)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged paths.' }
    foreach ($name in ($historyPaths + $staged)) {
        if ($name -and $name -notin $files) { throw "Unexpected file: $name. Review manually before upload." }
    }
    foreach ($name in $localFiles) {
        $destination = Join-Path $repo ('BD_GENERATOR/' + $name)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $destination
    }
    Invoke-Git (@('-C', $repo, 'add', '--') + $files)
    Invoke-Git @('-C', $repo, 'diff', '--cached', '--stat')
    if ($PrepareOnly) {
        Write-Host 'Prepared source-only staging. No commit, network access or push was performed.'
        exit 0
    }
    Write-Host "Repository: $RepositoryUrl"
    Write-Host "Branch: $branch"
    Write-Host 'Only BD_GENERATOR source, tests and documentation will be committed. Review the list above.'
    if ((Read-Host 'Type UPLOAD to commit and push') -cne 'UPLOAD') {
        Write-Host 'Cancelled. Local staging remains available.'
        exit 0
    }
    $authorName = (& git -C $repo config user.name)
    $authorEmail = (& git -C $repo config user.email)
    if (-not $authorName) {
        $authorName = Read-Host 'Commit author name (saved only in this upload repository)'
        if (-not $authorName.Trim()) { throw 'Author name is required. Nothing was pushed.' }
        Invoke-Git @('-C', $repo, 'config', 'user.name', $authorName.Trim())
    }
    if (-not $authorEmail) {
        $authorEmail = Read-Host 'Commit email (public in history; GitHub noreply email recommended)'
        if (-not $authorEmail.Trim()) { throw 'Author email is required. Nothing was pushed.' }
        Invoke-Git @('-C', $repo, 'config', 'user.email', $authorEmail.Trim())
    }
    Invoke-Git @('-C', $repo, 'remote', 'set-url', 'origin', $RepositoryUrl)
    & git -C $repo diff --cached --quiet
    if ($LASTEXITCODE -eq 1) {
        Invoke-Git @('-C', $repo, 'commit', '-m', 'feat: Alpha 0.9 Word comparison from Alpha 0.4 packages')
    } elseif ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged changes.' }
    Invoke-Git @('-C', $repo, 'push', '-u', 'origin', $branch)
    Write-Host 'Upload completed. Analyzer history retained; no force push used.'
} catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
