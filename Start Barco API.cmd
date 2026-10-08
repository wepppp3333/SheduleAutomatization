@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start_automation_server.ps1"
echo.
echo Barco API window finished. Press any key to close.
pause >nul
