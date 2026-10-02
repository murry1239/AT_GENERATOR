@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0UPLOAD_TO_GITHUB.ps1"
if errorlevel 1 echo Upload was not completed. No force push was attempted.
pause
endlocal
