@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Process-PmH1.ps1"
if errorlevel 1 (
  echo Rebuild failed.
  pause
  exit /b 1
)
echo.
echo Starting phone dashboard. Keep this window open.
echo On your phone, open the Phone Wi-Fi URL shown next.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Start-MobileServer.ps1"
pause
