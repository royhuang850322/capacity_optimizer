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
call %PY_CMD% -c "import ortools, pandas, openpyxl, click, colorama"
if errorlevel 1 goto :deps_missing

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
