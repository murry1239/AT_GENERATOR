@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD="

where py >nul 2>nul
if not errorlevel 1 (
  py -3.11 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>nul
  if not errorlevel 1 set "PY_CMD=py -3.11"
)

if not defined PY_CMD (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"
  )
)

if not defined PY_CMD goto no_python

echo [1/3] Creating virtual environment...
%PY_CMD% -m venv ".venv"
if errorlevel 1 goto failed

set "VENV_PY=%CD%\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" goto failed

echo [2/3] Updating pip...
"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 goto failed

echo [3/3] Installing required package...
"%VENV_PY%" -m pip install -r "requirements.txt"
if errorlevel 1 goto failed

echo.
echo Installation completed successfully.
echo Double-click RUN.bat to start the program.
goto finish

:no_python
echo.
echo Python 3.11 or later was not found.
echo Install 64-bit Python and enable the Python Launcher, then run this file again.
goto finish

:failed
echo.
echo Installation failed.
echo Run DIAGNOSE.bat and save its output for troubleshooting.

:finish
echo.
pause
endlocal
