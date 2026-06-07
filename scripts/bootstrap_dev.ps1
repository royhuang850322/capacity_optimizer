[CmdletBinding()]
param(
    [string]$Python = "python",
    [switch]$SkipVenv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $RepoRoot

if (-not $SkipVenv) {
    $VenvPath = Join-Path $RepoRoot ".venv"
    $VenvPython = Join-Path $VenvPath "Scripts\python.exe"

    if (-not (Test-Path -LiteralPath $VenvPython)) {
        Write-Host "Creating virtual environment: $VenvPath"
        & $Python -m venv $VenvPath
    }

    $PythonExe = $VenvPython
} else {
    $PythonExe = $Python
}

Write-Host "Using Python: $PythonExe"
& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r requirements.txt
& $PythonExe -m pip install -r requirements-dev.txt

Write-Host ""
Write-Host "Development dependencies are installed."
Write-Host "Run the release preflight check with:"
Write-Host "  .\scripts\release_preflight.ps1"
