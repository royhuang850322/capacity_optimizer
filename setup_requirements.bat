@echo off
setlocal
title Capacity Optimizer - Dependency Setup
cd /d "%~dp0"

echo.
echo Capacity Optimizer - Dependency Setup
echo ------------------------------------
echo Project folder:
echo   %~dp0
echo.

if not exist "%~dp0requirements.txt" (
    echo ERROR: requirements.txt not found.
    goto :fail
)

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

if not defined PY_CMD (
    echo ERROR: Python was not found on this computer.
    echo Install Python first, then run this file again.
    goto :fail
)

echo Using Python command: %PY_CMD%
call %PY_CMD% --version
if errorlevel 1 goto :fail

echo.
echo Clearing old pip temp/cache files...
call %PY_CMD% -m pip cache purge >nul 2>nul
powershell -NoProfile -Command "Remove-Item \"$env:LOCALAPPDATA\Temp\pip-*\" -Recurse -Force -ErrorAction SilentlyContinue" >nul 2>nul

echo.
echo Installing required packages...
call %PY_CMD% -m pip install --user --no-cache-dir -r requirements.txt
if errorlevel 1 goto :retry

goto :verify

:retry
echo.
echo First install attempt failed. Retrying ortools separately...
call %PY_CMD% -m pip install --user --no-cache-dir ortools
if errorlevel 1 goto :fail

call %PY_CMD% -m pip install --user --no-cache-dir -r requirements.txt
if errorlevel 1 goto :fail

:verify
echo.
echo Verifying Python imports...
call %PY_CMD% -c "import ortools, pandas, openpyxl, click, colorama; print('Dependency check: OK')"
if errorlevel 1 goto :fail

echo.
echo Setup completed successfully.
echo You can now open Capacity_Optimizer_Control.xlsx and run the tool.
goto :end

:fail
echo.
echo Setup did not complete.
echo Check network, antivirus, or Python installation, then run this file again.
exit /b 1

:end
echo.
pause
exit /b 0
