@echo off
rem ================================================================
rem  QuickBooks TimeWarp(R) by Our System Administrator  —  Launcher
rem
rem  USAGE:
 rem   1. Double-click this file to open the GUI.
rem   2. Drag one or more .QBW files onto this icon to pre-load them.
rem ================================================================

cd /d "%~dp0"

rem --- Find Python (check PATH first, then common installs) ---
where python >nul 2>&1
if %errorlevel%==0 (
    set PYTHON_CMD=python
    goto :launch
)

where python3 >nul 2>&1
if %errorlevel%==0 (
    set PYTHON_CMD=python3
    goto :launch
)

rem Check common install locations
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
    "C:\Python39\python.exe"
) do (
    if exist %%~P (
        set PYTHON_CMD=%%~P
        goto :launch
    )
)

echo.
echo  ERROR: Python not found on this system.
echo.
echo  Please install Python 3.9+ from https://python.org
echo  Make sure to check "Add Python to PATH" during install.
echo.
pause
exit /b 1

:launch
echo  Starting QuickBooks TimeWarp by Our System Administrator...
echo.

rem --- Pre-launch cleanup: kill stale processes so nothing is locked ---
rem    This runs BEFORE the GUI starts, so there's no race condition.
echo  Cleaning up stale processes...
taskkill /F /IM QBW32PremierAccountant.exe >nul 2>&1
taskkill /F /IM QBWPremierAccountant.exe   >nul 2>&1
taskkill /F /IM qbw32.exe                  >nul 2>&1
taskkill /F /IM qbw.exe                    >nul 2>&1
taskkill /F /IM qbupdate.exe               >nul 2>&1
rem Kill any orphaned TimeWarp Python processes (but NOT python.exe
rem globally — other scripts may be running). We target the exact
rem script name so only TimeWarp instances die.
wmic process where "commandline like '%%qb_downgrade_gui%%'" call terminate >nul 2>&1
wmic process where "commandline like '%%qb_automation%%'" call terminate >nul 2>&1

rem Give Windows a moment to release file locks after kills
timeout /t 3 /nobreak >nul

rem Clean stale working directories (locked files from crashed runs)
if exist "%~dp0..\Working\source" (
    echo  Removing stale working\source...
    rmdir /s /q "%~dp0..\Working\source" >nul 2>&1
)
if exist "%~dp0..\Working\Export" (
    echo  Removing stale working\Export...
    rmdir /s /q "%~dp0..\Working\Export" >nul 2>&1
)
echo  Clean.
echo.

echo  Detecting drive layout...
%PYTHON_CMD% "%~dp0drive_layout.py"
echo.
%PYTHON_CMD% "%~dp0qb_downgrade_gui.py" %*

if %errorlevel% neq 0 (
    echo.
    echo  TimeWarp exited with an error. See logs\ folder for details.
    echo.
    pause
)
