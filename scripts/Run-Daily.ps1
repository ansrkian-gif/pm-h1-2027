# Rebuild dashboard from the OWS Excel extract in this folder.
$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Process-PmH1.ps1")
