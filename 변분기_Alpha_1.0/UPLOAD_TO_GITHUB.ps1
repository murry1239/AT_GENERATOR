param([switch]$PrepareOnly, [string]$RepositoryUrl = 'https://github.com/murry1239/AT_GENERATOR.git')
$ErrorActionPreference = 'Stop'
function Invoke-Git {
    param([string[]]$Arguments)
    & git @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Git failed: $($Arguments -join ' ')" }
}
try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Install Git for Windows first.' }
    if ($RepositoryUrl -notmatch '^https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') { throw 'Use a GitHub HTTPS repository URL without a token.' }
    $suiteRoot = Split-Path -Parent $PSScriptRoot
    $analyzerFolder = '변분기_Alpha_1.0'; $generatorFolder = '변대생기_Alpha_1.0'
    $analyzerFiles = @(
        '.gitignore', 'analyzer.py', 'app.py', 'pair_analysis.py', 'pair_runner.py', 'word_capture.py', 'comparison_scope.py',
        'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md', 'INSTALL.bat', 'RUN.bat', 'DIAGNOSE.bat', 'TEST.bat',
        'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1', 'tests/test_pair_analysis.py', 'tests/test_alpha04.py',
        'tests/test_alpha05.py', 'tools/validate_sample.py', 'tools/build_release.py', 'tools/build_suite_release.py'
    )
    $generatorFiles = @(
        '.gitignore', 'app.py', 'generator.py', 'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md',
        'INSTALL.bat', 'RUN.bat', 'TEST.bat', 'DIAGNOSE.bat', 'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1',
        'tests/test_generator.py', 'tools/build_release.py'
    )
    $sourcePaths = @($analyzerFiles | ForEach-Object { "$analyzerFolder/$_" }) + @($generatorFiles | ForEach-Object { "$generatorFolder/$_" })
    foreach ($name in $sourcePaths) {
        if (-not (Test-Path -LiteralPath (Join-Path $suiteRoot $name) -PathType Leaf)) { throw "Missing suite source: $name. Keep both Alpha 1.0 folders next to each other." }
    }
    $branch = 'codex/alpha-1.0'; $zipName = 'AT_GENERATOR_Alpha_1.0.zip'
    $allowed = @($sourcePaths) + $zipName
    $workName = if ($PrepareOnly) { '.github-preview' } else { '.github-upload' }
    $repo = Join-Path $PSScriptRoot $workName
    if (-not (Test-Path -LiteralPath (Join-Path $repo '.git'))) {
        if (Test-Path -LiteralPath $repo) { throw 'Upload directory exists without Git metadata; inspect it first.' }
        if ($PrepareOnly) { Invoke-Git @('init', '-b', $branch, $repo) }
        else {
            Invoke-Git @('clone', $RepositoryUrl, $repo)
            & git -C $repo show-ref --verify --quiet "refs/remotes/origin/$branch"
            if ($LASTEXITCODE -eq 0) { Invoke-Git @('-C', $repo, 'switch', '-c', $branch, '--track', "origin/$branch") }
            else { Invoke-Git @('-C', $repo, 'switch', '-c', $branch) }
        }
    }
    $actualRoot = (& git -C $repo rev-parse --show-toplevel)
    if ($LASTEXITCODE -ne 0 -or [IO.Path]::GetFullPath($actualRoot) -ne [IO.Path]::GetFullPath($repo)) { throw 'Upload directory must be an independent Git repository.' }
    $actualBranch = (& git -C $repo symbolic-ref --short HEAD)
    if ($LASTEXITCODE -ne 0 -or $actualBranch -ne $branch) { throw "Expected branch: $branch" }
    & git -C $repo diff --quiet
    if ($LASTEXITCODE -ne 0) { throw 'Upload copy has unstaged edits. Preserve/review them before rerunning.' }
    $staged = @(& git -C $repo -c core.quotepath=false diff --cached --name-only)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged paths.' }
    foreach ($name in $staged) { if ($name -and $name -notin $allowed) { throw "Unexpected staged file: $name. Review manually." } }
    foreach ($name in $sourcePaths) {
        $destination = Join-Path $repo $name
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
        Copy-Item -LiteralPath (Join-Path $suiteRoot $name) -Destination $destination
    }
    # Explicit whitelist only: no recursive collection of documents or environments.
    Add-Type -AssemblyName System.IO.Compression
    $zipPath = Join-Path $repo $zipName
    $zipStream = [IO.File]::Open($zipPath, [IO.FileMode]::Create)
    $archive = [IO.Compression.ZipArchive]::new($zipStream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $sourcePaths) {
            $entry = $archive.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $entry.Open(); $sourceStream = [IO.File]::OpenRead((Join-Path $suiteRoot $name))
            try { $sourceStream.CopyTo($entryStream) }
            finally { $sourceStream.Dispose(); $entryStream.Dispose() }
        }
    } finally { $archive.Dispose(); $zipStream.Dispose() }
    Invoke-Git (@('-C', $repo, 'add', '--') + $sourcePaths)
    Invoke-Git @('-C', $repo, 'add', '-f', '--', $zipName)
    Invoke-Git @('-C', $repo, 'diff', '--cached', '--stat')
    if ($PrepareOnly) { Write-Host 'Prepared both source folders and the combined ZIP. No network, commit, or push was performed.'; exit 0 }
    Write-Host "Repository: $RepositoryUrl"
    Write-Host "Branch: $branch"
    if ((Read-Host 'Review the listed files; type UPLOAD to commit and push') -cne 'UPLOAD') { Write-Host 'Cancelled. Local staging is preserved.'; exit 0 }
    $authorName = (& git -C $repo config user.name); $authorEmail = (& git -C $repo config user.email)
    if (-not $authorName) {
        $authorName = Read-Host 'Commit author name'
        if (-not $authorName.Trim()) { throw 'Commit author name is required.' }
        Invoke-Git @('-C', $repo, 'config', 'user.name', $authorName.Trim())
    }
    if (-not $authorEmail) {
        $authorEmail = Read-Host 'Commit email (GitHub noreply email recommended)'
        if (-not $authorEmail.Trim()) { throw 'Commit email is required.' }
        Invoke-Git @('-C', $repo, 'config', 'user.email', $authorEmail.Trim())
    }
    Invoke-Git @('-C', $repo, 'remote', 'set-url', 'origin', $RepositoryUrl)
    & git -C $repo diff --cached --quiet
    if ($LASTEXITCODE -eq 1) { Invoke-Git @('-C', $repo, 'commit', '-m', 'release: Alpha 1.0 analyzer and comparison generator suite') }
    elseif ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged changes.' }
    Invoke-Git @('-C', $repo, 'push', '-u', 'origin', $branch)
    Write-Host 'Uploaded both folders and ZIP. Open a pull request for later updates to main.'
} catch { Write-Error $_ -ErrorAction Continue; exit 1 }
