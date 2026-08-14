# Opens TeleOWS so you can export Task Work, then rebuilds the dashboard when the Excel file changes.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Extract = Join-Path $Root "PM H1 2027.xlsx"
$url = "https://106d-sg.teleows.com/"
Write-Host "Opening $url"
Start-Process $url
Write-Host "Login -> left menu Task -> Query task (ooredoo)"
Write-Host "Complete Time From 2026-08-01  To now"
Write-Host "Task Type PM. Status Closed + Completed."
Write-Host "Task ID / Title / Site FME: _OGK Active and Passive"
Write-Host "Search -> Extract. Save over: $Extract"
Write-Host "Watching for a new save..."
$start = (Get-Item -LiteralPath $Extract).LastWriteTime
for ($i = 0; $i -lt 180; $i++) {
    Start-Sleep -Seconds 5
    $now = (Get-Item -LiteralPath $Extract).LastWriteTime
    if ($now -gt $start) {
        Write-Host "New extract detected at $now"
        & (Join-Path $PSScriptRoot "Process-PmH1.ps1")
        Write-Host "Dashboard updated. Run Start-Mobile.bat to view on the phone."
        exit 0
    }
}
Write-Host "No new Excel save detected. Dashboard not rebuilt."
exit 1
