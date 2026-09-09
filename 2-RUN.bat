@echo off
setlocal
title TR Plus Ultra - Test Run
pushd "%~dp0"

echo.
echo  ===============================================
echo    TR Plus Ultra  -  Test Run (no build needed)
echo  ===============================================
echo.

set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (python --version >nul 2>&1 && set "PY=python")

if not defined PY (
  echo  [X] Python not found on this PC.
  echo.
  echo      Install from https://www.python.org/downloads/
  echo      and tick "Add python.exe to PATH" during setup.
  goto theend
)

echo  Python found: %PY%
%PY% --version
echo.

echo  Checking libraries...
%PY% -c "import playwright, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo  Installing libraries, first time may take 2-3 minutes...
  echo.
  %PY% -m pip install --upgrade playwright openpyxl
  if errorlevel 1 (
    echo.
    echo  [X] Library install failed. Check your internet and try again.
    goto theend
  )
) else (
  echo  Libraries OK.
)

echo.
echo  Starting program...
echo.
%PY% tr_launcher.py
echo.
echo  Program closed.

:theend
echo.
popd
pause
