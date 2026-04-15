[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Get-SharedModuleConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $lines = Get-Content -LiteralPath $ConfigPath
    $section = $null
    $vbaDirectory = $null
    $modules = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {
        $trimmed = $line.Trim()

        if ($trimmed -match '^\[(.+)\]$') {
            $section = $Matches[1]
            continue
        }

        if ($section -eq 'project' -and -not $vbaDirectory -and $trimmed -match '^vba_directory\s*=\s*"(.+)"$') {
            $vbaDirectory = $Matches[1]
            continue
        }

        if ($section -eq 'shared' -and $trimmed -match '^modules\s*=\s*\[(.+)\]$') {
            $raw = $Matches[1]
            $raw -split ',' | ForEach-Object {
                $entry = $_.Trim().Trim('"')
                if ($entry -ne '') { $modules.Add($entry) }
            }
            continue
        }
    }

    if (-not $vbaDirectory) {
        throw "Could not resolve [project].vba_directory from $ConfigPath"
    }

    return [PSCustomObject]@{
        ConfigName   = [System.IO.Path]::GetFileName($ConfigPath)
        VbaDirectory = $vbaDirectory
        Modules      = @($modules)
    }
}

$configDir  = Join-Path $PWD "config"
$sharedDir  = Join-Path $PWD "src\shared"
$configFiles = Get-ChildItem -LiteralPath $configDir -Filter *.toml | Sort-Object Name

if ($configFiles.Count -eq 0) {
    throw "No .toml files found in config/. Nothing to sync."
}

$totalCopied = 0

foreach ($configFile in $configFiles) {
    $cfg = Get-SharedModuleConfig -ConfigPath $configFile.FullName

    if ($cfg.Modules.Count -eq 0) {
        Write-Host "$($cfg.ConfigName): no [shared].modules declared, skipping."
        continue
    }

    Write-Host "$($cfg.ConfigName): syncing $($cfg.Modules.Count) module(s) into $($cfg.VbaDirectory)"

    foreach ($module in $cfg.Modules) {
        $src  = Join-Path $PWD (Join-Path "src\shared" $module)
        $dest = Join-Path $PWD (Join-Path $cfg.VbaDirectory $module)

        if (-not (Test-Path -LiteralPath $src)) {
            throw "Shared module not found: src/shared/$module (declared in $($cfg.ConfigName)). Stopping."
        }

        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        Copy-Item -LiteralPath $src -Destination $dest -Force
        Write-Host "  Copied: src/shared/$module -> $($cfg.VbaDirectory)/$module"
        $totalCopied++
    }
}

Write-Host ""
Write-Host "Sync complete. $totalCopied file(s) copied."
Write-Host "Review changes with: git diff src/"
Write-Host "Commit the changes before running import-vba."
