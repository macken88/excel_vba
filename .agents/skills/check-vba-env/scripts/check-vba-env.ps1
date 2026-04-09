[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Get-ConfigWorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $line = Select-String -LiteralPath $ConfigPath -Pattern '^file\s*=\s*"(.+)"$' | Select-Object -First 1
    if (-not $line) {
        return $null
    }

    return $line.Matches[0].Groups[1].Value
}

function Get-VbaEditVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonPath
    )

    $stdout = [System.IO.Path]::GetTempFileName()
    $stderr = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process -FilePath $PythonPath `
            -ArgumentList @("-m", "pip", "show", "vba-edit") `
            -WorkingDirectory $PWD `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $stdout `
            -RedirectStandardError $stderr

        if ($process.ExitCode -ne 0) {
            return "not-available"
        }

        $versionLine = Get-Content -LiteralPath $stdout | Where-Object { $_ -like "Version:*" } | Select-Object -First 1
        if ($versionLine) {
            return $versionLine.Replace("Version:", "").Trim()
        }

        return "installed"
    }
    finally {
        Remove-Item -LiteralPath $stdout, $stderr -ErrorAction SilentlyContinue
    }
}

$venvPython = Join-Path $PWD ".venv-vba-tools\Scripts\python.exe"
$gitDir = Join-Path $PWD ".git"

Write-Host "Repo root: $PWD"
Write-Host "Git initialized: $(Test-Path -LiteralPath $gitDir)"
Write-Host "Venv python present: $(Test-Path -LiteralPath $venvPython)"

if (Test-Path -LiteralPath $venvPython) {
    Write-Host "vba-edit version: $(Get-VbaEditVersion -PythonPath $venvPython)"
}

Get-ChildItem -LiteralPath (Join-Path $PWD "config") -Filter *.toml | Sort-Object Name | ForEach-Object {
    $relativeWorkbook = Get-ConfigWorkbookPath -ConfigPath $_.FullName
    $workbookPath = if ($relativeWorkbook) { Join-Path $PWD $relativeWorkbook } else { $null }
    $exists = if ($workbookPath) { Test-Path -LiteralPath $workbookPath } else { $false }

    Write-Host "Config: $($_.Name)"
    Write-Host "  workbook: $relativeWorkbook"
    Write-Host "  workbook_exists: $exists"
}
