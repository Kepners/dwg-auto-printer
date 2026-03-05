param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectPath = Join-Path $repoRoot "src\DwgAutoPrinter.App\DwgAutoPrinter.App.csproj"
$distRoot = Join-Path $repoRoot "dist"
$legacyRoot = Join-Path $distRoot "releases"
$indexPath = Join-Path $distRoot "index.json"

if (-not (Test-Path $projectPath)) {
    throw "Project file not found: $projectPath"
}

New-Item -ItemType Directory -Path $distRoot -Force | Out-Null

$existingReleaseDirs = @()
$existingReleaseDirs += Get-ChildItem -Path $distRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^rev-(\d{4})(?:_.*)?$' }
if (Test-Path $legacyRoot) {
    $existingReleaseDirs += Get-ChildItem -Path $legacyRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^rev-(\d{4})(?:_.*)?$' }
}

$maxRevision = 0
foreach ($dir in $existingReleaseDirs) {
    $match = [regex]::Match($dir.Name, '^rev-(\d{4})(?:_.*)?$')
    if ($match.Success) {
        $value = [int]$match.Groups[1].Value
        if ($value -gt $maxRevision) {
            $maxRevision = $value
        }
    }
}

$nextRevision = $maxRevision + 1
$revisionLabel = ("{0:D4}" -f $nextRevision)
$releaseName = "rev-$revisionLabel"
$releasePath = Join-Path $distRoot $releaseName

Write-Host "Publishing revision $revisionLabel to $releasePath"

dotnet publish $projectPath `
    -c $Configuration `
    -r $Runtime `
    --self-contained false `
    /p:PublishSingleFile=true `
    -o $releasePath

if (-not (Test-Path $releasePath)) {
    throw "Publish failed. Release output missing: $releasePath"
}

$exePath = Join-Path $releasePath "DwgAutoPrinter.App.exe"
$lspPath = Join-Path $releasePath "smart-revision-update.lsp"

if (-not (Test-Path $exePath)) {
    throw "Expected EXE not found in release output: $exePath"
}

if (-not (Test-Path $lspPath)) {
    throw "Expected LISP not found in release output: $lspPath"
}

$releaseEntry = [ordered]@{
    revision = $revisionLabel
    timestamp = (Get-Date).ToString("s")
    runtime = $Runtime
    configuration = $Configuration
    releaseFolder = $releaseName
    exe = "$releaseName/DwgAutoPrinter.App.exe"
    lsp = "$releaseName/smart-revision-update.lsp"
}

$indexData = @()
if (Test-Path $indexPath) {
    $raw = Get-Content -Path $indexPath -Raw
    if ($raw.Trim().Length -gt 0) {
        $parsed = $raw | ConvertFrom-Json
        if ($parsed -is [System.Array]) {
            $indexData = @($parsed)
        } elseif ($null -ne $parsed) {
            $indexData = @($parsed)
        }
    }
}

$indexData += [pscustomobject]$releaseEntry
ConvertTo-Json -InputObject @($indexData) -Depth 4 | Set-Content -Path $indexPath -Encoding UTF8

Write-Host ""
Write-Host "Dist release created:"
Write-Host "  Revision: $revisionLabel"
Write-Host "  Folder:   $releasePath"
Write-Host "  EXE:      $exePath"
Write-Host "  LISP:     $lspPath"
Write-Host "  Index:    $indexPath"
