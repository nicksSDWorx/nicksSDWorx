@echo off
REM Run Dashboard AI Worx from source (no build required).
setlocal
cd /d "%~dp0"

python -c "import webview" 2>nul
if errorlevel 1 (
  echo Installing pywebview...
  python -m pip install pywebview || goto :error
)

python dashboard.py
exit /b %errorlevel%

:error
echo Install FAILED. Make sure Python 3.10+ is on your PATH.
pause
exit /b 1
