@echo off
REM Wekelijks voortgangsrapport - dubbelklik om te draaien.
REM Werkt vanuit de projectmap waarin dit .bat-bestand staat.

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
)

python monitor.py
set EXITCODE=%ERRORLEVEL%

if not "%EXITCODE%"=="0" (
    echo.
    echo ============================================================
    echo Er ging iets mis bij het genereren van het rapport.
    echo Lees de melding hierboven en raadpleeg README.md bij twijfel.
    echo ============================================================
    pause
)

exit /b %EXITCODE%
