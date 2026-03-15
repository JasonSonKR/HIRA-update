[CmdletBinding()]
param(
    [string]$TaskName = "HIRA Material Rolling Sync",
    [datetime]$At = (Get-Date "06:00")
)

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "run_hira_material_sync.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`" -Mode rolling -Scheduled"
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Description "Daily trigger at 06:00 that runs HIRA rolling sync only on the 5th, 15th, and 25th." -Force | Out-Null
Write-Output "Registered scheduled task: $TaskName"
