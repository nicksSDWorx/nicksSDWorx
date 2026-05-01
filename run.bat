@echo off
REM Double-click to play. Requires Python 3 (with tkinter, which is bundled
REM by default on the python.org Windows installer).
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 cosmic_catcher.py
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python cosmic_catcher.py
    goto :end
)

echo.
echo Python 3 was not found on this PC.
echo Install it from https://www.python.org/downloads/  (tick "Add Python to PATH"),
echo then double-click run.bat again.
echo.
pause

:end
endlocal
