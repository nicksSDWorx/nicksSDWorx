@echo off
REM Double-click to play Payday Panic. Requires Python 3 (tkinter ships
REM with the python.org Windows installer).
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 payday_panic.py
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python payday_panic.py
    goto :end
)

echo.
echo Python 3 was not found on this PC.
echo Install it from https://www.python.org/downloads/  (tick "Add Python to PATH"),
echo then double-click run_payday.bat again.
echo.
pause

:end
endlocal
