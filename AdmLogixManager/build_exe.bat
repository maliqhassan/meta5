@echo off
REM Builds a stand-alone program in dist\OfficeManager\OfficeManager.exe
REM (no Python needed on the computer that runs it - copy the whole folder).
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --noconsole --onedir --name OfficeManager app.py
echo.
echo Done. Open dist\OfficeManager and double-click OfficeManager.exe
echo Your data will be saved in dist\OfficeManager\data\office.db
pause
