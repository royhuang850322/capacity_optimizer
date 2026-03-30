@echo off
setlocal
title Capacity Optimizer - Run
cd /d "%~dp0\.."

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

set "TEMPLATE=%CD%\Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
set "ACTIVE_LICENSE=%CD%\licenses\active\license.json"
set "LEGACY_LICENSE=%CD%\license.json"

echo.
echo Capacity Optimizer - Excel Workflow
echo ----------------------------------
echo Control workbook:
echo   %TEMPLATE%
echo.

if not defined PY_CMD (
    echo Python was not found on this computer.
    echo Install Python first, then run runtime\setup_requirements.bat.
    goto :end
)

if not exist "%TEMPLATE%" (
    echo Control workbook not found. Creating a fresh workbook...
    call %PY_CMD% -m app.create_template
    if errorlevel 1 goto :end
)

echo Checking required Python packages...
call %PY_CMD% -c "import ortools, pandas, openpyxl, click, colorama, cryptography"
if errorlevel 1 goto :deps_missing

if not exist "%ACTIVE_LICENSE%" if not exist "%LEGACY_LICENSE%" (
    echo.
    echo License file not found.
    echo Checked:
    echo   %ACTIVE_LICENSE%
    echo   %LEGACY_LICENSE%
    echo If you already have a trial or unbound license, copy license.json into licenses\active\.
    echo Otherwise run runtime\get_machine_fingerprint.bat, send the generated file from licenses\requests\ to RSCP,
    echo and place the returned machine-locked license.json into licenses\active\.
    goto :end
)

echo Running optimizer...
call %PY_CMD% -m app.main --input-template "%TEMPLATE%"
if errorlevel 1 goto :end

if exist "%CD%\output" (
    echo.
    echo Opening output folder...
    start "" explorer "%CD%\output"
)

goto :end

:deps_missing
echo.
echo Required Python packages are missing or incomplete.
echo Run runtime\setup_requirements.bat first, then try Run Optimizer again.

:end
echo.
pause
