@echo off
REM Bakes cosmic_catcher.py into a single-file Windows .exe using PyInstaller.
REM Output ends up in dist\CosmicCatcher.exe and is fully standalone -
REM the player does NOT need Python installed to run that .exe.
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

echo Building CosmicCatcher.exe (this takes ~30s the first time)...
%PY% -m PyInstaller --onefile --windowed --name CosmicCatcher cosmic_catcher.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done.  Your game is at:  dist\CosmicCatcher.exe
echo Double-click it to play - no Python needed on the target machine.
pause
endlocal
