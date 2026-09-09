@echo off
chcp 65001 >nul
echo ============================================
echo   TR Plus Ultra - Build Launcher EXE
echo ============================================
echo.

echo [1/3] ติดตั้งไลบรารีที่ต้องใช้...
pip install --upgrade pyinstaller playwright openpyxl
if errorlevel 1 goto fail

echo.
echo [2/3] กำลัง build...
pyinstaller --noconfirm --onefile --windowed ^
  --name "TR Plus Ultra" ^
  --collect-all playwright ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.worksheet.datavalidation ^
  --hidden-import openpyxl.utils ^
  tr_launcher.py
if errorlevel 1 goto fail

echo.
echo [3/3] เสร็จแล้ว
echo ไฟล์อยู่ที่  dist\TR Plus Ultra.exe
echo แจกไฟล์นี้ให้ทีมได้เลย ทำครั้งเดียวพอ
echo.
pause
exit /b 0

:fail
echo.
echo *** Build ไม่สำเร็จ ***
pause
exit /b 1
