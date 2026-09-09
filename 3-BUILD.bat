@echo off
setlocal
title TR Plus Ultra - Build EXE
pushd "%~dp0"

echo.
echo  ===============================================
echo    TR Plus Ultra  -  Build EXE
echo  ===============================================
echo.

if not exist "tr_launcher.py" (
  echo  [X] tr_launcher.py not found.
  echo      Put this .bat in the same folder as tr_launcher.py
  goto theend
)

findstr /C:"GITHUB_USER = 'CHANGE-ME'" tr_launcher.py >nul 2>&1
if not errorlevel 1 (
  echo  [X] GITHUB_USER is still CHANGE-ME in tr_launcher.py
  echo      Edit that line first, then run this again.
  goto theend
)

echo  [1/3] Looking for Python...
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (python --version >nul 2>&1 && set "PY=python")

if not defined PY (
  echo.
  echo  [X] Python not found.
  echo      Install from https://www.python.org/downloads/
  echo      and tick "Add python.exe to PATH" during setup.
  goto theend
)
echo        Found: %PY%
%PY% --version
echo.

echo  [2/3] Installing build tools (first time may take a few minutes)...
echo.
%PY% -m pip install --upgrade pip
%PY% -m pip install --upgrade pyinstaller playwright openpyxl
if errorlevel 1 (
  echo.
  echo  [X] Install failed. Check internet and try again.
  goto theend
)
echo.

echo  [3/3] Building EXE, this takes 3-5 minutes. Do not close this window.
echo.
%PY% -m PyInstaller --noconfirm --onefile --windowed --name "TR Plus Ultra" --collect-all playwright --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import openpyxl.worksheet.datavalidation tr_launcher.py
if errorlevel 1 (
  echo.
  echo  [X] Build failed. Please screenshot this window.
  goto theend
)

echo.
if exist "dist\TR Plus Ultra.exe" (
  echo  ===============================================
  echo    SUCCESS
  echo    File: dist\TR Plus Ultra.exe
  echo    Share this file with your team.
  echo  ===============================================
  echo.
  start "" "%~dp0dist"
) else (
  echo  [X] dist\TR Plus Ultra.exe not found. Please screenshot this window.
)

:theend
echo.
popd
pause
