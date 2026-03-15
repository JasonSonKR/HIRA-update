[CmdletBinding()]
param(
    [string]$TaskName = "HIRA Material Sync",
    [string]$DayOfWeek = "Monday",
    [datetime]$At = (Get-Date "10:15")
)

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "run_hira_material_sync.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $At
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Description "Weekly HIRA material sync" -Force | Out-Null
Write-Output "Registered scheduled task: $TaskName"
