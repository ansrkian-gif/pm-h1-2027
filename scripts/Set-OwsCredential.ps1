# Store TeleOWS login in Windows Credential Manager. Password is not saved in this folder.
$ErrorActionPreference = "Stop"
$user = Read-Host "TeleOWS username"
$sec = Read-Host "TeleOWS password" -AsSecureString
$plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
)
cmdkey /generic:TeleOWS-PMH1 /user:$user /pass:$plain | Out-Null
Write-Host "Saved Windows credential target TeleOWS-PMH1"
Write-Host "OWS export is still a manual save until the live pull is connected."
