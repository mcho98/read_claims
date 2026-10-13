@echo off
REM Build ClaimFiller.exe on Windows. Needs Python once, on the build machine only.
REM Works from any folder: it always builds inside python-app (the folder this file is in).
cd /d "%~dp0"
pip install openpyxl playwright pyinstaller || exit /b 1
pyinstaller --onefile --windowed --name ClaimFiller --collect-all playwright app.py || exit /b 1
copy /Y config.json dist\config.json
echo Done. Give the user the files in python-app\dist (ClaimFiller.exe and config.json).
