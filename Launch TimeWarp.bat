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
%PYTHON_CMD% "%~dp0qb_downgrade_gui.py" %*

if %errorlevel% neq 0 (
    echo.
    echo  TimeWarp exited with an error. See logs\ folder for details.
    echo.
    pause
)
