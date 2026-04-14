@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Capacity Optimizer - Python Check
cd /d "%~dp0\.."

echo.
echo Capacity Optimizer - Python Environment Check
echo --------------------------------------------
echo Project folder:
echo   %CD%
echo.

set "PY_CMD="
set "PY_VERSION="
set "PY_SOURCE="
set "STORE_ALIAS_MSG="

call :probe_python "python" "PATH"
if defined PY_CMD goto :success

call :probe_python "py -3" "PY_LAUNCHER"
if defined PY_CMD goto :success

echo No usable Python command was found.
echo.
echo Searching common installation folders...
call :find_installed_python
if defined FOUND_PY goto :configure_found

echo.
echo Python does not appear to be installed correctly on this computer.
echo.
echo Recommended fix:
echo 1. Install Python 3.12 or newer from https://www.python.org/downloads/windows/
echo 2. During installation, check "Add python.exe to PATH"
echo 3. Re-open Command Prompt or PowerShell
echo 4. Run runtime\check_python_setup.bat again
echo.
if defined STORE_ALIAS_MSG (
    echo Additional note:
    echo Windows is currently showing the Microsoft Store Python alias.
    echo If Python is already installed, go to:
    echo   Settings ^> Apps ^> Advanced app settings ^> App execution aliases
    echo Then turn off python.exe and python3.exe aliases.
)
goto :end

:configure_found
echo Found Python installation:
echo   %FOUND_PY%
echo.
echo Configuring PATH for the current user...
set "FOUND_DIR=%FOUND_PY%\.."
for %%I in ("%FOUND_DIR%") do set "FOUND_DIR=%%~fI"
set "FOUND_SCRIPTS=%FOUND_DIR%\Scripts"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$py='%FOUND_DIR%'; $scripts='%FOUND_SCRIPTS%';" ^
  "$current=[Environment]::GetEnvironmentVariable('Path','User');" ^
  "$parts=@(); if($current){ $parts=$current -split ';' | Where-Object { $_ -and $_.Trim() -ne '' } };" ^
  "foreach($item in @($py,$scripts)){ if($item -and -not ($parts -contains $item)){ $parts += $item } };" ^
  "[Environment]::SetEnvironmentVariable('Path', ($parts -join ';'), 'User')"
if errorlevel 1 (
    echo Failed to update PATH automatically.
    echo Add these folders to the user PATH manually:
    echo   %FOUND_DIR%
    echo   %FOUND_SCRIPTS%
    goto :end
)

set "PATH=%FOUND_DIR%;%FOUND_SCRIPTS%;%PATH%"
call :probe_python "python" "CONFIGURED_PATH"
if defined PY_CMD goto :success

echo PATH was updated, but this window still could not resolve python.
echo Close this window, open a new Command Prompt, and run runtime\check_python_setup.bat again.
goto :end

:success
echo Python is ready to use.
echo.
echo Command:
echo   %PY_CMD%
echo Version:
echo   %PY_VERSION%
echo Source:
echo   %PY_SOURCE%
echo.
echo Next step:
echo 1. Run runtime\setup_requirements.bat
echo 2. If needed, run runtime\get_machine_fingerprint.bat
goto :end

:probe_python
set "CANDIDATE=%~1"
set "SOURCE=%~2"
set "CMD_OUT="
for /f "usebackq delims=" %%I in (`cmd /d /c "%CANDIDATE% --version 2^>^&1"`) do (
    if not defined CMD_OUT set "CMD_OUT=%%I"
)
if not defined CMD_OUT goto :eof

echo %CMD_OUT% | findstr /i /c:"Python was not found" >nul
if not errorlevel 1 (
    set "STORE_ALIAS_MSG=1"
    goto :eof
)

echo %CMD_OUT% | findstr /b /i /c:"Python " >nul
if errorlevel 1 goto :eof

set "PY_CMD=%CANDIDATE%"
set "PY_VERSION=%CMD_OUT%"
set "PY_SOURCE=%SOURCE%"
goto :eof

:find_installed_python
set "FOUND_PY="
for /f "usebackq delims=" %%I in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$roots=@($env:LOCALAPPDATA + '\Programs\Python', $env:ProgramFiles, ${env:ProgramFiles(x86)});" ^
  "$hits=@();" ^
  "foreach($root in $roots){ if(Test-Path $root){ $hits += Get-ChildItem -Path $root -Filter python.exe -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match 'Python\d+' } } }" ^
  "$best=$hits | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName;" ^
  "if($best){ $best }"`) do (
    set "FOUND_PY=%%I"
)
goto :eof

:end
echo.
pause
exit /b 0
