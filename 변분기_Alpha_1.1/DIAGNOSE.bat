@echo off
setlocal
cd /d "%~dp0"

echo === Change Analyzer Alpha 1.1 diagnostics ===
echo Working directory: %CD%
echo.

echo [Python launcher]
where py 2>nul
py --version 2>nul
echo.

echo [Python command]
where python 2>nul
python --version 2>nul
echo.

echo [Virtual environment]
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" --version
  ".venv\Scripts\python.exe" -m pip show pywin32
  ".venv\Scripts\python.exe" -m pip show openpyxl
  ".venv\Scripts\python.exe" -m pip show Pillow
  ".venv\Scripts\python.exe" -m pip show lxml
) else (
  echo NOT FOUND: .venv\Scripts\python.exe
)
echo.

echo [Microsoft Word]
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2>nul
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve 2>nul
echo.

echo Diagnostics finished.
pause
endlocal
