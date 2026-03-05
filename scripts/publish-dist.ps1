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
$tempPublishPath = Join-Path $distRoot "_publish-temp"

if (-not (Test-Path $projectPath)) {
    throw "Project file not found: $projectPath"
}

New-Item -ItemType Directory -Path $distRoot -Force | Out-Null

function Get-MaxRevision {
    param(
        [string]$DistPath,
        [string]$LegacyPath,
        [string]$IndexFile
    )

    $max = 0

    # New format: files in dist root, e.g. DwgAutoPrinter.App.rev-0007.exe
    $rootFiles = Get-ChildItem -Path $DistPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^DwgAutoPrinter\.App\.rev-(\d{4})\.exe$' }
    foreach ($file in $rootFiles) {
        $v = [int]([regex]::Match($file.Name, '^DwgAutoPrinter\.App\.rev-(\d{4})\.exe$').Groups[1].Value)
        if ($v -gt $max) { $max = $v }
    }

    # Legacy folder format: rev-0007 or rev-0007_...
    $legacyDirs = @()
    $legacyDirs += Get-ChildItem -Path $DistPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^rev-(\d{4})(?:_.*)?$' }
    if (Test-Path $LegacyPath) {
        $legacyDirs += Get-ChildItem -Path $LegacyPath -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^rev-(\d{4})(?:_.*)?$' }
    }
    foreach ($dir in $legacyDirs) {
        $v = [int]([regex]::Match($dir.Name, '^rev-(\d{4})(?:_.*)?$').Groups[1].Value)
        if ($v -gt $max) { $max = $v }
    }

    # Index history fallback
    if (Test-Path $IndexFile) {
        $raw = Get-Content -Path $IndexFile -Raw
        if ($raw.Trim().Length -gt 0) {
            $parsed = $raw | ConvertFrom-Json
            $items = @()
            if ($parsed -is [System.Array]) {
                $items = @($parsed)
            } elseif ($null -ne $parsed) {
                $items = @($parsed)
            }

            foreach ($item in $items) {
                $r = [string]$item.revision
                if ($r -match '^\d{4}$') {
                    $v = [int]$r
                    if ($v -gt $max) { $max = $v }
                }
            }
        }
    }

    return $max
}

$maxRevision = Get-MaxRevision -DistPath $distRoot -LegacyPath $legacyRoot -IndexFile $indexPath
$nextRevision = $maxRevision + 1
$revisionLabel = ("{0:D4}" -f $nextRevision)

Write-Host "Publishing revision $revisionLabel to dist root (no revision folder)"

if (Test-Path $tempPublishPath) {
    Remove-Item -Path $tempPublishPath -Recurse -Force
}
New-Item -ItemType Directory -Path $tempPublishPath -Force | Out-Null

dotnet publish $projectPath `
    -c $Configuration `
    -r $Runtime `
    --self-contained false `
    /p:PublishSingleFile=true `
    -o $tempPublishPath

$tempExe = Join-Path $tempPublishPath "DwgAutoPrinter.App.exe"
$tempLsp = Join-Path $tempPublishPath "smart-revision-update.lsp"
$tempPdb = Join-Path $tempPublishPath "DwgAutoPrinter.App.pdb"

if (-not (Test-Path $tempExe)) {
    throw "Expected EXE not found in temp publish output: $tempExe"
}
if (-not (Test-Path $tempLsp)) {
    throw "Expected LISP not found in temp publish output: $tempLsp"
}

$latestExe = Join-Path $distRoot "DwgAutoPrinter.App.exe"
$latestLsp = Join-Path $distRoot "smart-revision-update.lsp"
$latestPdb = Join-Path $distRoot "DwgAutoPrinter.App.pdb"

$revExe = Join-Path $distRoot ("DwgAutoPrinter.App.rev-" + $revisionLabel + ".exe")
$revLsp = Join-Path $distRoot ("smart-revision-update.rev-" + $revisionLabel + ".lsp")
$revPdb = Join-Path $distRoot ("DwgAutoPrinter.App.rev-" + $revisionLabel + ".pdb")

Copy-Item -Path $tempExe -Destination $latestExe -Force
Copy-Item -Path $tempExe -Destination $revExe -Force
Copy-Item -Path $tempLsp -Destination $latestLsp -Force
Copy-Item -Path $tempLsp -Destination $revLsp -Force

if (Test-Path $tempPdb) {
    Copy-Item -Path $tempPdb -Destination $latestPdb -Force
    Copy-Item -Path $tempPdb -Destination $revPdb -Force
}

$releaseEntry = [ordered]@{
    revision = $revisionLabel
    timestamp = (Get-Date).ToString("s")
    runtime = $Runtime
    configuration = $Configuration
    exe = "DwgAutoPrinter.App.rev-$revisionLabel.exe"
    lsp = "smart-revision-update.rev-$revisionLabel.lsp"
    latestExe = "DwgAutoPrinter.App.exe"
    latestLsp = "smart-revision-update.lsp"
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

if (Test-Path $tempPublishPath) {
    Remove-Item -Path $tempPublishPath -Recurse -Force
}

Write-Host ""
Write-Host "Dist release created:"
Write-Host "  Revision:    $revisionLabel"
Write-Host "  Latest EXE:  $latestExe"
Write-Host "  Latest LISP: $latestLsp"
Write-Host "  Rev EXE:     $revExe"
Write-Host "  Rev LISP:    $revLsp"
Write-Host "  Index:       $indexPath"
