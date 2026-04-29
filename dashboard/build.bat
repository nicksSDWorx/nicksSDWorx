@echo off
REM Build a standalone Windows .exe for Dashboard AI Worx.
REM No external runtime deps: the app uses only Python stdlib at runtime.
REM PyInstaller is the only build-time requirement.

setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python not found on PATH. Install from https://www.python.org/
  pause
  exit /b 1
)
python --version

echo.
echo [1/3] Installing PyInstaller...
python -m pip install --upgrade pip >nul
python -m pip install pyinstaller
if errorlevel 1 goto :deperror

echo.
echo [2/3] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist DashboardAIWorx.spec del /q DashboardAIWorx.spec

echo.
echo [3/3] Building DashboardAIWorx.exe...
python -m PyInstaller ^
  --name DashboardAIWorx ^
  --noconsole ^
  --onefile ^
  --add-data "ui.html;." ^
  --add-data "tool_window.html;." ^
  dashboard.py
if errorlevel 1 goto :builderror

echo.
echo ============================================================
echo  Done! Executable: %cd%\dist\DashboardAIWorx.exe
echo ============================================================
echo.
pause
exit /b 0

:deperror
echo Dependency install FAILED. See error above.
pause
exit /b 1

:builderror
echo PyInstaller build FAILED. See error above.
pause
exit /b 1
