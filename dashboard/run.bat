@echo off
REM Run Dashboard AI Worx from source (no deps required — pure stdlib).
setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python not found on PATH. Install from https://www.python.org/
  pause
  exit /b 1
)

python dashboard.py
exit /b %errorlevel%
