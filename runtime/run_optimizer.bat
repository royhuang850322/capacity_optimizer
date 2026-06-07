@echo off
setlocal
title Capacity Optimizer - Launcher
cd /d "%~dp0\.."

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

echo.
echo Capacity Optimizer - Desktop Launcher
echo -------------------------------------
echo The Excel control workbook UI has been retired.
echo Use CapacityOptimizerLauncher.pyw or the packaged CapacityOptimizer.exe.
echo.

if not defined PY_CMD (
    echo Python was not found on this computer.
    echo Run runtime\check_python_setup.bat first.
    goto :end
)

if not exist "%CD%\CapacityOptimizerLauncher.pyw" (
    echo CapacityOptimizerLauncher.pyw was not found in this package.
    goto :end
)

echo Starting launcher...
start "" %PY_CMD% "%CD%\CapacityOptimizerLauncher.pyw"

:end
echo.
pause
