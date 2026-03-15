[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string[]]$InputPath,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$pythonCandidates = @(
    (Join-Path $repoRoot ".venv\Scripts\python.exe"),
    (Join-Path (Split-Path -Parent $repoRoot) ".venv\Scripts\python.exe")
)
$pythonExe = $pythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$scriptPath = Join-Path $PSScriptRoot "process_hira_mhtml_xls.py"

if (-not $pythonExe) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        throw "Python executable was not found. Install Python or create a .venv folder in the repo root."
    }
    $pythonExe = $pythonCommand.Source
}

$arguments = @($scriptPath, "--config", $ConfigPath)
if ($InputPath) {
    $arguments += "--input"
    $arguments += $InputPath
}
if ($Force) {
    $arguments += "--force"
}

& $pythonExe $arguments
exit $LASTEXITCODE
