@echo off
REM ============================================================
REM  Build script for Stamkaart Word naar Excel Converter
REM  Run this on a Windows machine with Python 3.10+ installed.
REM ============================================================

echo === Stamkaart Converter - Build Script ===
echo.

REM Install dependencies
echo [1/2] Installing dependencies...
pip install python-docx openpyxl pyinstaller
if %ERRORLEVEL% neq 0 (
    echo FOUT: Dependencies konden niet worden geinstalleerd.
    pause
    exit /b 1
)

echo.
echo [2/2] Building .exe with PyInstaller...
pyinstaller --onefile --windowed --name "StamkaartConverter" app.py
if %ERRORLEVEL% neq 0 (
    echo FOUT: PyInstaller build mislukt.
    pause
    exit /b 1
)

echo.
echo === Build voltooid! ===
echo Het .exe bestand staat in: dist\StamkaartConverter.exe
echo.
pause
