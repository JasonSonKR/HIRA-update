$ErrorActionPreference = "Stop"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 (Join-Path $scriptDirectory "pull_hira_results_from_github.py")
    exit $LASTEXITCODE
}

& python (Join-Path $scriptDirectory "pull_hira_results_from_github.py")
exit $LASTEXITCODE
