@echo off
setlocal
title Capacity Optimizer - Run
cd /d "%~dp0"

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

set "TEMPLATE=%~dp0Tooling Control Panel\Capacity_Optimizer_Control.xlsx"

echo.
echo Capacity Optimizer - Excel Workflow
echo ----------------------------------
echo Control workbook:
echo   %TEMPLATE%
echo.

if not defined PY_CMD (
    echo Python was not found on this computer.
    echo Install Python first, then run setup_requirements.bat.
    goto :end
)

if not exist "%TEMPLATE%" (
    echo Control workbook not found. Creating a fresh workbook...
    call %PY_CMD% create_template.py
    if errorlevel 1 goto :end
)

echo Checking required Python packages...
call %PY_CMD% -c "import ortools, pandas, openpyxl, click, colorama, cryptography"
if errorlevel 1 goto :deps_missing

if not exist "%~dp0license.json" (
    echo.
    echo License file not found: %~dp0license.json
    echo If you already have a trial or unbound license, copy license.json into the project root.
    echo Otherwise run get_machine_fingerprint.bat, send machine_fingerprint.json to RSCP,
    echo and place the returned machine-locked license.json in the project root.
    goto :end
)

echo Running optimizer...
call %PY_CMD% main.py --input-template "%TEMPLATE%"
if errorlevel 1 goto :end

if exist "%~dp0output" (
    echo.
    echo Opening output folder...
    start "" explorer "%~dp0output"
)

goto :end

:deps_missing
echo.
echo Required Python packages are missing or incomplete.
echo Run setup_requirements.bat first, then try Run Optimizer again.

:end
echo.
pause
