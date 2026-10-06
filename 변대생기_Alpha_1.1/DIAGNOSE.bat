@echo off
setlocal
cd /d "%~dp0"
echo Change Table Generator Alpha 1.1
echo Working directory: %CD%
where py 2>nul
where python 2>nul
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" --version
  ".venv\Scripts\python.exe" -m pip show python-docx Pillow
  ".venv\Scripts\python.exe" -c "import tkinter, docx; print('Tk and DOCX engine OK')"
) else (
  echo Virtual environment missing. Run INSTALL.bat.
)
pause
endlocal

