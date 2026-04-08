[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Resolve-PythonExe {
    $candidates = @(
        "C:\Users\sim_m\AppData\Local\Programs\Python\Python312\python.exe",
        "C:\Users\sim_m\AppData\Local\Programs\Python\Python311\python.exe",
        "C:\Program Files\Python312\python.exe",
        "C:\Program Files\Python311\python.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "Python executable not found in expected Windows locations."
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

        if (Test-Path -LiteralPath $stdout) {
            Get-Content -LiteralPath $stdout
        }

        if ($process.ExitCode -ne 0) {
            if (Test-Path -LiteralPath $stderr) {
                Get-Content -LiteralPath $stderr
            }
            throw "Command failed with exit code $($process.ExitCode): $FilePath $($ArgumentList -join ' ')"
        }
    }
    finally {
        Remove-Item -LiteralPath $stdout, $stderr -ErrorAction SilentlyContinue
    }
}

$pythonExe = Resolve-PythonExe
$venvDir = Join-Path $PWD ".venv-vba-tools"
$venvPython = Join-Path $venvDir "Scripts\python.exe"

Write-Host "Using Python: $pythonExe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Creating venv: $venvDir"
    Invoke-External -FilePath $pythonExe -ArgumentList @("-m", "venv", ".venv-vba-tools")
}
else {
    Write-Host "Reusing existing venv: $venvDir"
}

Invoke-External -FilePath $venvPython -ArgumentList @("-m", "pip", "install", "--upgrade", "pip")
Invoke-External -FilePath $venvPython -ArgumentList @("-m", "pip", "install", "-r", "requirements-vba-tools.txt")
Invoke-External -FilePath $venvPython -ArgumentList @("-m", "pip", "show", "vba-edit")

Write-Host "Bootstrap completed."
