@echo off
setlocal
title Capacity Optimizer - License Generator
cd /d "%~dp0\.."

set "PY_CMD="
set "PYW_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"
where pythonw >nul 2>nul
if not errorlevel 1 set "PYW_CMD=pythonw"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

echo.
echo Chemical Capacity Optimizer - License Generator
echo -----------------------------------------------
echo Project folder:
echo   %CD%
echo.

if not defined PY_CMD (
    echo Python was not found on this computer.
    echo Install Python first, then run runtime\setup_requirements.bat.
    goto :end
)

echo Checking required Python packages...
call %PY_CMD% -c "import tkinter, cryptography"
if errorlevel 1 (
    echo.
    echo Required Python packages are missing or incomplete.
    echo Run runtime\setup_requirements.bat first, then try again.
    goto :end
)

if defined PYW_CMD (
    start "" %PYW_CMD% license_admin\license_tools\license_generator_ui.py
    exit /b 0
)

call %PY_CMD% license_admin\license_tools\license_generator_ui.py
if errorlevel 1 (
    echo.
    echo The license generator did not start successfully.
)

:end
echo.
pause
exit /b 0
