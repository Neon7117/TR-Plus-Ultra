@echo off
setlocal enabledelayedexpansion
title TR Plus Ultra - System Check
pushd "%~dp0"
set "R=%~dp0report.txt"

echo.
echo  Checking your system, please wait...
echo.

call :report > "%R%" 2>&1

echo  Done.
echo  Saved to: report.txt
echo.
start "" notepad "%R%"
popd
pause
exit /b

:report
echo ============ TR Plus Ultra - System Report ============
echo.
echo [folder]
echo %~dp0
echo.
echo [where python]
where python
echo.
echo [where py]
where py
echo.
echo [where pip]
where pip
echo.
echo [py -0p : installed python versions]
py -0p
echo.
echo [py -3 --version]
py -3 --version
echo.
echo [python --version]
python --version
echo.
echo [py -3 -m pip --version]
py -3 -m pip --version
echo.
echo [import tkinter]
py -3 -c "import tkinter; print('tkinter OK')"
echo.
echo [import playwright]
py -3 -c "import playwright; print('playwright OK')"
echo.
echo [import openpyxl]
py -3 -c "import openpyxl; print('openpyxl OK')"
echo.
echo [LOCALAPPDATA Programs Python]
dir /b "%LOCALAPPDATA%\Programs\Python"
echo.
echo [Microsoft Store stub]
if exist "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" (echo FOUND store stub) else (echo no store stub)
echo.
echo [Chrome]
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (echo Chrome in ProgramFiles) else (echo not in ProgramFiles)
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (echo Chrome in ProgramFiles x86) else (echo not in ProgramFiles x86)
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (echo Chrome in LocalAppData) else (echo not in LocalAppData)
echo.
echo [PATH]
echo %PATH%
echo.
echo [files here]
dir /b
echo.
echo ============ end of report ============
exit /b
