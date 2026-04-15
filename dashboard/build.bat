@echo off
REM Build a standalone Windows .exe for Dashboard AI Worx.
REM Usage: double-click this file, or run "build.bat" from a terminal.

setlocal EnableDelayedExpansion
cd /d "%~dp0"

call :find_python
if errorlevel 1 goto :pyerror

echo Using Python: !PY_CMD!
!PY_CMD! --version

echo.
echo [1/3] Installing build dependencies (wheels only)...
!PY_CMD! -m pip install --upgrade pip >nul
REM --only-binary=pythonnet prevents pip from trying to compile the
REM legacy pythonnet 2.5.x from source (which requires NuGet and fails
REM on modern Python). Pinning to >=3.0 forces the modern package that
REM ships wheels.
!PY_CMD! -m pip install --only-binary=pythonnet "pythonnet>=3.0" pywebview pyinstaller
if errorlevel 1 goto :deperror

echo.
echo [2/3] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist DashboardAIWorx.spec del /q DashboardAIWorx.spec

echo.
echo [3/3] Building DashboardAIWorx.exe...
!PY_CMD! -m PyInstaller ^
  --name DashboardAIWorx ^
  --noconsole ^
  --onefile ^
  --add-data "ui.html;." ^
  dashboard.py
if errorlevel 1 goto :builderror

echo.
echo ============================================================
echo  Done! Executable: %cd%\dist\DashboardAIWorx.exe
echo ============================================================
echo.
pause
exit /b 0

REM ------------------------------------------------------------
REM Helpers
REM ------------------------------------------------------------

:find_python
REM Try the py launcher for 3.13 down to 3.10 first (pywebview's
REM native deps have wheels for these). Then fall back to "python"
REM if it's within that range.
set PY_CMD=
for %%V in (3.13 3.12 3.11 3.10) do (
  if not defined PY_CMD (
    py -%%V -c "import sys" >nul 2>&1
    if not errorlevel 1 set PY_CMD=py -%%V
  )
)
if not defined PY_CMD (
  python -c "import sys; sys.exit(0 if (3,10) <= sys.version_info[:2] <= (3,13) else 1)" >nul 2>&1
  if not errorlevel 1 set PY_CMD=python
)
if not defined PY_CMD exit /b 1
exit /b 0

:pyerror
echo.
echo ============================================================
echo  No compatible Python interpreter found.
echo.
echo  pywebview needs Python 3.10 - 3.13 on Windows, because its
echo  dependency "pythonnet" has no pre-built wheels for Python
echo  3.14 yet, and building it from source requires Visual Studio
echo  + NuGet and usually fails.
echo.
echo  Fix: install Python 3.12 or 3.13 from https://www.python.org/
echo       (leave "Add to PATH" on). Then just run this bat again
echo       - it will pick up the new interpreter automatically.
echo ============================================================
echo.
pause
exit /b 1

:deperror
echo.
echo Installing dependencies FAILED. See error above.
pause
exit /b 1

:builderror
echo.
echo PyInstaller build FAILED. See error above.
pause
exit /b 1
