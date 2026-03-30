@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

if not defined PY_CMD (
    echo Python was not found on this computer.
    pause
    exit /b 1
)

echo.
echo Chemical Capacity Optimizer - Machine Fingerprint
echo -------------------------------------------------
call %PY_CMD% machine_fingerprint.py --out "%~dp0machine_fingerprint.json"
if errorlevel 1 (
    echo.
    echo Failed to generate machine fingerprint.
    pause
    exit /b 1
)

echo.
echo machine_fingerprint.json has been created in:
echo   %~dp0machine_fingerprint.json
echo Please send this file to RSCP to request a machine-locked license.
pause
exit /b 0
