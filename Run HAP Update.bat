@echo off
setlocal
cd /d "%~dp0"
title HAP Database Update

echo ============================================================
echo   HAP DATABASE UPDATE
echo ============================================================
echo.

rem --- find Python ---
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )

if not defined PY (
    echo Python is not installed on this PC.
    echo I'll open the download page. Install it, then TICK the box
    echo   "Add python.exe to PATH", finish, and run this again.
    start "" "https://www.python.org/downloads/"
    echo.
    pause
    exit /b 1
)

echo Using Python: %PY%
echo.
echo Checking required components (first run only, ~1 min)...
%PY% -m pip install --quiet --disable-pip-version-check --upgrade pip >nul 2>&1
%PY% -m pip install --quiet --disable-pip-version-check pandas openpyxl xlrd requests
if errorlevel 1 (
    echo.
    echo Could not install the required components automatically.
    echo Please connect to the internet and try again.
    echo.
    pause
    exit /b 1
)

echo.
%PY% "%~dp0hap_update.py"

endlocal
