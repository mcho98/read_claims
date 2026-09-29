@echo off
REM Build ClaimFiller.exe on Windows. Needs Python once, on the build machine only.
pip install openpyxl playwright pyinstaller || exit /b 1
pyinstaller --onefile --windowed --name ClaimFiller --collect-all playwright app.py || exit /b 1
copy /Y config.json dist\config.json
echo Done. Give the user the files in the dist folder (ClaimFiller.exe and config.json).
