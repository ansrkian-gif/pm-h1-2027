# Register a daily 06:30 rebuild (Kuwait morning). Runs only when this Windows account is logged on.
$ErrorActionPreference = "Stop"
$script = Join-Path $PSScriptRoot "Run-Daily.ps1"
$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`""
$tn = "PM H1 2027 Daily Dashboard"
schtasks /Create /F /TN $tn /SC DAILY /ST 06:30 /TR $action /RL LIMITED
Write-Host "Scheduled task created: $tn at 06:30 daily"
Write-Host "It rebuilds from PM H1 2027.xlsx. Replace that file after the OWS export."
