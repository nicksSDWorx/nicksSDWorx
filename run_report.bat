@echo off
REM Wekelijks voortgangsrapport -- dubbelklik om te draaien.
REM Bij de eerste run wordt automatisch een virtuele omgeving aangemaakt
REM en worden de benodigde Python-packages geinstalleerd.

setlocal
cd /d "%~dp0"

REM === Stap 1: is Python beschikbaar? ===
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ============================================================
    echo Python is niet gevonden op deze computer.
    echo.
    echo Doe het volgende:
    echo   1. Ga naar https://www.python.org/downloads/
    echo   2. Download Python 3.10 of nieuwer.
    echo   3. Tijdens de installatie: vink aan "Add Python to PATH".
    echo   4. Dubbelklik daarna opnieuw op run_report.bat.
    echo ============================================================
    echo.
    pause
    exit /b 1
)

REM === Stap 2: virtuele omgeving + packages installeren (eenmalig) ===
if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo Eerste keer draaien -- benodigde packages worden nu geinstalleerd.
    echo Dit duurt 1 a 2 minuten en gebeurt maar eenmalig.
    echo.
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo ============================================================
        echo Kon de virtuele omgeving niet aanmaken.
        echo Controleer of Python correct is geinstalleerd.
        echo ============================================================
        echo.
        pause
        exit /b 1
    )
    call ".venv\Scripts\activate.bat"
    python -m pip install --quiet --disable-pip-version-check --upgrade pip
    python -m pip install --quiet --disable-pip-version-check -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ============================================================
        echo Installatie van packages mislukt.
        echo Mogelijke oorzaken:
        echo   - Geen internet op dit moment ^(alleen tijdens installatie nodig^).
        echo   - Bedrijfs-proxy blokkeert pip.
        echo Neem contact op met IT als dit blijft mislukken.
        echo ============================================================
        echo.
        pause
        exit /b 1
    )
    echo.
    echo Installatie geslaagd. Het rapport wordt nu gegenereerd...
    echo.
) else (
    call ".venv\Scripts\activate.bat"
)

REM === Stap 3: rapport genereren ===
python monitor.py
set EXITCODE=%ERRORLEVEL%

if not "%EXITCODE%"=="0" (
    echo.
    echo ============================================================
    echo Er ging iets mis bij het genereren van het rapport.
    echo Lees de melding hierboven en raadpleeg README.md bij twijfel.
    echo ============================================================
    echo.
    pause
)

exit /b %EXITCODE%
