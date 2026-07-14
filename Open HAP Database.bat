@echo off
setlocal
cd /d "%~dp0"
title HAP Database
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY ( python --version >nul 2>&1 && set "PY=python" )
if not defined PY (
  echo Python isn't installed yet. Opening the download page.
  echo Install it, TICK "Add python.exe to PATH", then open this again.
  start "" "https://www.python.org/downloads/"
  pause & exit /b 1
)
%PY% -m pip install --quiet --disable-pip-version-check pandas openpyxl xlrd requests >nul 2>&1
echo Starting HAP Database... your browser will open in a moment.
echo (Keep this small window open while you use it.)
%PY% "%~dp0hap_server.py"
endlocal
