[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Get-ProjectConfigValues {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $lines = Get-Content -LiteralPath $ConfigPath
    $inProjectSection = $false
    $file = $null
    $vbaDirectory = $null

    foreach ($line in $lines) {
        $trimmed = $line.Trim()

        if ($trimmed -match '^\[(.+)\]$') {
            $inProjectSection = ($Matches[1] -eq 'project')
            continue
        }

        if (-not $inProjectSection) {
            continue
        }

        if (-not $file -and $trimmed -match '^file\s*=\s*"(.+)"$') {
            $file = $Matches[1]
            continue
        }

        if (-not $vbaDirectory -and $trimmed -match '^vba_directory\s*=\s*"(.+)"$') {
            $vbaDirectory = $Matches[1]
            continue
        }
    }

    if (-not $file -or -not $vbaDirectory) {
        throw "Could not resolve [project] file/vba_directory from $ConfigPath"
    }

    return @{
        File = $file
        VbaDirectory = $vbaDirectory
    }
}

function Invoke-External {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process -FilePath $FilePath `
            -ArgumentList $ArgumentList `
            -WorkingDirectory $PWD `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdout `
            -RedirectStandardError $stderr

        $output = @()

        if (Test-Path -LiteralPath $stdout) {
            $output += Get-Content -LiteralPath $stdout
        }

        if (Test-Path -LiteralPath $stderr) {
            $output += Get-Content -LiteralPath $stderr
        }

        return @{
            ExitCode = $process.ExitCode
            Output = $output
        }
    }
    finally {
        Remove-Item -LiteralPath $stdout, $stderr -ErrorAction SilentlyContinue
    }
}

$excelVbaExe = Join-Path $PWD ".venv-vba-tools\Scripts\excel-vba.exe"

if (-not (Test-Path -LiteralPath $excelVbaExe)) {
    throw "excel-vba.exe not found in .venv-vba-tools. Run bootstrap-vba-env first."
}

$checkResult = Invoke-External -FilePath $excelVbaExe -ArgumentList @("check")
$checkResult.Output | ForEach-Object { Write-Host $_ }

$checkText = ($checkResult.Output -join "`n")
if ($checkText -match "seems to be disabled") {
    throw "Excel VBA Trust Access is not enabled. Import stopped before modifying any workbook."
}

Get-ChildItem -LiteralPath (Join-Path $PWD "config") -Filter *.toml | Sort-Object Name | ForEach-Object {
    Write-Host "Importing with config: $($_.Name)"
    $projectConfig = Get-ProjectConfigValues -ConfigPath $_.FullName
    $importResult = Invoke-External -FilePath $excelVbaExe -ArgumentList @(
        "import",
        "--file", $projectConfig.File,
        "--vba-directory", $projectConfig.VbaDirectory
    )
    $importResult.Output | ForEach-Object { Write-Host $_ }

    if ($importResult.ExitCode -ne 0) {
        throw "Import failed for config $($_.Name)"
    }
}

Write-Host "Import completed for all configured workbooks."
