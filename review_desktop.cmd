@echo off
setlocal
cd /d "%~dp0"
if not exist ".conda-env\python.exe" (
  echo Missing project Python environment. Run: python -m venv .conda-env
  echo Then run: .conda-env\python.exe -m pip install -e ".[desktop]"
  pause
  exit /b 1
)
".conda-env\python.exe" review_desktop.py
if errorlevel 1 pause
