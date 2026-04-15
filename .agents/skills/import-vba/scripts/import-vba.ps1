[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$BookFiles
)

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

    return [PSCustomObject]@{
        ConfigPath = $ConfigPath
        ConfigName = [System.IO.Path]::GetFileName($ConfigPath)
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

function Get-ConfiguredWorkbooks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigDirectory
    )

    return Get-ChildItem -LiteralPath $ConfigDirectory -Filter *.toml |
        Sort-Object Name |
        ForEach-Object { Get-ProjectConfigValues -ConfigPath $_.FullName }
}

function Find-WorkbookCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedBookFile,

        [Parameter(Mandatory = $true)]
        [object[]]$ConfiguredWorkbooks
    )

    $requestedLeaf = [System.IO.Path]::GetFileName($RequestedBookFile)
    $candidates = New-Object System.Collections.Generic.List[string]

    foreach ($workbook in $ConfiguredWorkbooks) {
        if ($workbook.File -ieq $RequestedBookFile) {
            if (-not $candidates.Contains($workbook.File)) {
                $candidates.Add($workbook.File)
            }
            continue
        }

        $configuredLeaf = [System.IO.Path]::GetFileName($workbook.File)
        if ($configuredLeaf -ieq $requestedLeaf -or $workbook.File -like "*$RequestedBookFile*") {
            if (-not $candidates.Contains($workbook.File)) {
                $candidates.Add($workbook.File)
            }
        }
    }

    return $candidates
}

function Resolve-RequestedWorkbooks {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$RequestedBookFiles,

        [Parameter(Mandatory = $true)]
        [object[]]$ConfiguredWorkbooks
    )

    $resolved = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    foreach ($requestedBookFile in $RequestedBookFiles) {
        $exactMatch = @($ConfiguredWorkbooks | Where-Object { $_.File -ceq $requestedBookFile })

        if ($exactMatch.Count -eq 1) {
            if (-not $seen.ContainsKey($exactMatch[0].File)) {
                $resolved.Add($exactMatch[0])
                $seen[$exactMatch[0].File] = $true
            }
            continue
        }

        $candidates = Find-WorkbookCandidates -RequestedBookFile $requestedBookFile -ConfiguredWorkbooks $ConfiguredWorkbooks
        $message = "Requested workbook '$requestedBookFile' did not exactly match any [project].file entry in config/*.toml."

        if ($candidates.Count -gt 0) {
            $message += " Candidate values: $($candidates -join ', '). Rerun with one of the exact values."
        }
        else {
            $known = $ConfiguredWorkbooks | ForEach-Object { $_.File }
            $message += " Known values: $($known -join ', ')."
        }

        throw $message
    }

    return $resolved
}

$excelVbaExe = Join-Path $PWD ".venv-vba-tools\Scripts\excel-vba.exe"

if (-not (Test-Path -LiteralPath $excelVbaExe)) {
    throw "excel-vba.exe not found in .venv-vba-tools. Follow the setup steps in README.md first."
}

$checkResult = Invoke-External -FilePath $excelVbaExe -ArgumentList @("check")
$checkResult.Output | ForEach-Object { Write-Host $_ }

$checkText = ($checkResult.Output -join "`n")
if ($checkText -match "seems to be disabled") {
    throw "Excel VBA Trust Access is not enabled. Import stopped before modifying any workbook."
}

$configuredWorkbooks = Get-ConfiguredWorkbooks -ConfigDirectory (Join-Path $PWD "config")
$requestedWorkbooks = Resolve-RequestedWorkbooks -RequestedBookFiles $BookFiles -ConfiguredWorkbooks $configuredWorkbooks

if ($requestedWorkbooks.Count -eq 0) {
    throw "No workbook imports were resolved from the provided BookFiles arguments."
}

Write-Host "Resolved workbook targets:"
$requestedWorkbooks | ForEach-Object {
    Write-Host "  $($_.File) <= $($_.ConfigName)"
}

$requestedWorkbooks | ForEach-Object {
    Write-Host "Importing with config: $($_.ConfigName)"
    $importResult = Invoke-External -FilePath $excelVbaExe -ArgumentList @(
        "import",
        "--file", $_.File,
        "--vba-directory", $_.VbaDirectory
    )
    $importResult.Output | ForEach-Object { Write-Host $_ }

    if ($importResult.ExitCode -ne 0) {
        throw "Import failed for workbook $($_.File) with config $($_.ConfigName)"
    }
}

Write-Host "Import completed for all requested workbooks."
