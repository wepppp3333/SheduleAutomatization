@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_automation_autostart.ps1"
echo.
echo Press any key to close.
pause >nul
