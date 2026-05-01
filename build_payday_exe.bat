@echo off
REM Bakes payday_panic.py into a single-file Windows .exe via PyInstaller.
REM Output: dist\PaydayPanic.exe (fully standalone, no Python needed).
setlocal
cd /d "%~dp0"

set PY=py -3
where py >nul 2>nul
if not %errorlevel%==0 (
    where python >nul 2>nul
    if not %errorlevel%==0 (
        echo Python 3 is required to build the .exe. Install it from python.org first.
        pause
        exit /b 1
    )
    set PY=python
)

echo Installing / updating PyInstaller...
%PY% -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo Could not install PyInstaller.
    pause
    exit /b 1
)

echo Building PaydayPanic.exe (this takes ~30s the first time)...
%PY% -m PyInstaller --onefile --windowed --name PaydayPanic payday_panic.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done.  Your game is at:  dist\PaydayPanic.exe
pause
endlocal
