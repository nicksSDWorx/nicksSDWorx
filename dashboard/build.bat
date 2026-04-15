@echo off
REM Build a standalone Windows .exe for Dashboard AI Worx.
REM Usage: double-click this file, or run "build.bat" from a terminal.

setlocal
cd /d "%~dp0"

echo [1/3] Installing build dependencies...
python -m pip install --upgrade pip >nul
python -m pip install pywebview pyinstaller || goto :error

echo [2/3] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist DashboardAIWorx.spec del /q DashboardAIWorx.spec

echo [3/3] Building DashboardAIWorx.exe...
python -m PyInstaller ^
  --name DashboardAIWorx ^
  --noconsole ^
  --onefile ^
  --add-data "ui.html;." ^
  dashboard.py || goto :error

echo.
echo ============================================================
echo  Done! Executable: %cd%\dist\DashboardAIWorx.exe
echo ============================================================
echo.
pause
exit /b 0

:error
echo.
echo Build FAILED. Scroll up for the error.
pause
exit /b 1
