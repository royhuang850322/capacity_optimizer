@echo off
setlocal
cd /d "%~dp0\.."

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
call %PY_CMD% -m app.machine_fingerprint --out-dir "%CD%\licenses\requests"
if errorlevel 1 (
    echo.
    echo Failed to generate machine fingerprint.
    pause
    exit /b 1
)

echo.
echo A machine fingerprint request file has been created in:
echo   %CD%\licenses\requests
echo Please send that file to RSCP to request a machine-locked license.
pause
exit /b 0
