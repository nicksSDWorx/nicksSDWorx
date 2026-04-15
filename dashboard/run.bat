@echo off
REM Run Dashboard AI Worx from source (no build required).
setlocal EnableDelayedExpansion
cd /d "%~dp0"

call :find_python
if errorlevel 1 goto :pyerror

!PY_CMD! -c "import webview" 2>nul
if errorlevel 1 (
  echo Installing pywebview ^(wheels only^)...
  !PY_CMD! -m pip install --only-binary=pythonnet "pythonnet>=3.0" pywebview
  if errorlevel 1 goto :deperror
)

!PY_CMD! dashboard.py
exit /b %errorlevel%

:find_python
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
echo  Please install Python 3.12 or 3.13 from https://www.python.org/
echo  (Python 3.14 is not supported yet by pywebview's deps.)
echo ============================================================
pause
exit /b 1

:deperror
echo Dependency install FAILED. See error above.
pause
exit /b 1
