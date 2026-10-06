@echo off
setlocal
cd /d "%~dp0"

set "VENV_PY=%CD%\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" goto not_installed

"%VENV_PY%" "analyzer.py"
if errorlevel 1 goto run_failed
endlocal
exit /b 0

:not_installed
echo The virtual environment was not found.
echo Run INSTALL.bat first.
goto finish

:run_failed
echo.
echo The program stopped with an error.
echo Run DIAGNOSE.bat and save its output for troubleshooting.

:finish
echo.
pause
endlocal
exit /b 1
