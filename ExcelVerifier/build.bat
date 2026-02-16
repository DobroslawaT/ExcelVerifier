@echo off
cd /d "%~dp0"
echo.
echo ============================================================
echo ExcelVerifier Build Script
echo ============================================================
echo.
echo [Step 1/2] Converting icon PNG to ICO...
.\.venv\Scripts\python convert_icon.py

echo.
echo [Step 2/2] Building ExcelVerifier.exe...
echo This will take 2-3 minutes. Please wait...
echo.

.\.venv\Scripts\pyinstaller --onefile --windowed ^
  --name="ExcelVerifier" ^
  --icon=icon.ico ^
  --add-data="ExcelVerifier/ui;ui" ^
  --add-data="ExcelVerifier/core;core" ^
  --collect-all=openpyxl ^
  ExcelVerifier/main.py

echo.
if exist "dist\ExcelVerifier.exe" (
    echo ============================================================
    echo SUCCESS! ExcelVerifier.exe created!
    echo ============================================================
    echo.
    echo Location: dist\ExcelVerifier.exe
    echo.
    echo You can now distribute this .exe to users!
    echo Users can simply run it - no Python required!
    echo.
    pause
) else (
    echo Build may have failed. Check the output above.
    pause
)
