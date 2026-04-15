[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Get-DiffOutput {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$GitArgs
    )

    $result = & git @GitArgs 2>&1
    return @($result)
}

$unstagedLines = Get-DiffOutput -GitArgs @("diff", "HEAD", "--", "src/")
$stagedLines   = Get-DiffOutput -GitArgs @("diff", "--cached", "HEAD", "--", "src/")

$hasUnstaged = $unstagedLines.Count -gt 0
$hasStaged   = $stagedLines.Count -gt 0

if (-not $hasUnstaged -and -not $hasStaged) {
    Write-Host "No pending changes in src/. Nothing to review."
    exit 0
}

if ($hasStaged) {
    Write-Host "=== Staged changes in src/ ==="
    $stagedLines | ForEach-Object { Write-Host $_ }
    Write-Host ""
}

if ($hasUnstaged) {
    Write-Host "=== Unstaged changes in src/ ==="
    $unstagedLines | ForEach-Object { Write-Host $_ }
    Write-Host ""
}

# Collect changed files
$changedFiles = New-Object System.Collections.Generic.List[string]

foreach ($line in ($stagedLines + $unstagedLines)) {
    if ($line -match '^\+\+\+\s+b/(.+)$') {
        $path = $Matches[1]
        if (-not $changedFiles.Contains($path)) {
            $changedFiles.Add($path)
        }
    }
}

if ($changedFiles.Count -gt 0) {
    Write-Host "=== Changed files ==="
    $changedFiles | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

# Map changed files to workbook configs
$configDir = Join-Path $PWD "config"
$configFiles = Get-ChildItem -LiteralPath $configDir -Filter *.toml | Sort-Object Name

$affectedWorkbooks = New-Object System.Collections.Generic.List[string]

foreach ($configFile in $configFiles) {
    $lines = Get-Content -LiteralPath $configFile.FullName
    $vbaDirectory = $null

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^vba_directory\s*=\s*"(.+)"$') {
            $vbaDirectory = $Matches[1]
            break
        }
    }

    if (-not $vbaDirectory) { continue }

    $normalizedVbaDir = $vbaDirectory.Replace('\', '/')
    $hasMatch = $changedFiles | Where-Object { $_.Replace('\', '/') -like "$normalizedVbaDir/*" }

    if ($hasMatch) {
        $affectedWorkbooks.Add("$($configFile.Name) ($vbaDirectory)")
    }
}

if ($affectedWorkbooks.Count -gt 0) {
    Write-Host "=== Affected workbook configs ==="
    $affectedWorkbooks | ForEach-Object { Write-Host "  $_" }
    Write-Host ""
}

$sharedChanged = $changedFiles | Where-Object { $_.Replace('\', '/') -like "src/shared/*" }
if ($sharedChanged) {
    Write-Host "NOTE: Shared module(s) in src/shared/ have changed. Run sync-shared-modules before import-vba."
    Write-Host ""
}

Write-Host "Review complete. Confirm the diff is intentional before committing or running import-vba."
