[CmdletBinding()]
param(
    [ValidateSet("rolling", "backfill", "range")]
    [string]$Mode = "rolling",
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$StartMonth = "",
    [string]$EndMonth = "",
    [string[]]$CategoryCode,
    [switch]$Force,
    [switch]$Scheduled
)

$ErrorActionPreference = "Stop"

function Resolve-PythonExe {
    $current = Get-Item -LiteralPath $PSScriptRoot
    while ($current) {
        $candidate = Join-Path $current.FullName ".venv\Scripts\python.exe"
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
        $current = $current.Parent
    }

    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        return $pythonCommand.Source
    }

    throw "Python executable was not found. Install Python or create a .venv folder in this workspace."
}

if ($Scheduled) {
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    $allowedDays = @($config.sync.rolling_days_of_month)
    if ($allowedDays -notcontains (Get-Date).Day) {
        Write-Output "Skipping scheduled sync because today is not in the configured run days."
        exit 0
    }
}

$pythonExe = Resolve-PythonExe
$scriptPath = Join-Path $PSScriptRoot "download_hira_material.py"
$arguments = @($scriptPath, "--config", $ConfigPath, "--mode", $Mode, "--browser", "chrome", "--headless")

if ($StartMonth) {
    $arguments += @("--start-month", $StartMonth)
}
if ($EndMonth) {
    $arguments += @("--end-month", $EndMonth)
}
if ($CategoryCode) {
    foreach ($code in $CategoryCode) {
        $arguments += @("--category-code", $code)
    }
}
if ($Force) {
    $arguments += "--force"
}

& $pythonExe $arguments
exit $LASTEXITCODE
