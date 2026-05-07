@echo off
setlocal

REM QuickBooks Downgrade Tool build script
REM Run this on Windows in a terminal opened at project root.

python -m venv .venv
call .venv\Scripts\activate

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

if not exist dist mkdir dist
if not exist build mkdir build

pyinstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "QuickBooksDowngradeTool" ^
  --icon docs\app_icon.ico ^
  --add-data "sample_config.json;." ^
  qb_downgrade_gui.py

echo.
echo Build complete. Executable available at dist\QuickBooksDowngradeTool\QuickBooksDowngradeTool.exe
pause
