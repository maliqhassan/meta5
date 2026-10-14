@echo off
REM Starts Office Manager. First run installs the two required libraries.
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Download it from https://www.python.org/downloads/
  echo During setup tick "Add python.exe to PATH", then run this file again.
  pause
  exit /b
)
python -c "import openpyxl, reportlab" 2>nul || python -m pip install -r requirements.txt
start "" pythonw app.py
