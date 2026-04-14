param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = (Resolve-Path (Join-Path $ScriptDir "..")).Path
$SpecPath = Join-Path $ScriptDir "CapacityOptimizer.spec"
$BuildRoot = Join-Path $ProjectRoot "build\\pyinstaller"
$DistRoot = Join-Path $ProjectRoot "dist"

Write-Host "Chemical Capacity Optimizer - PyInstaller One-Folder Build"
Write-Host "Project root : $ProjectRoot"
Write-Host "Spec file    : $SpecPath"
Write-Host "Build root   : $BuildRoot"
Write-Host "Dist root    : $DistRoot"

Push-Location $ProjectRoot
try {
    python --version | Out-Null
} catch {
    Pop-Location
    throw "Python was not found on PATH. Install Python first, then rerun this script."
}

try {
    python -c "import PyInstaller" | Out-Null
} catch {
    Pop-Location
    throw "PyInstaller is not installed in the current Python environment. Run: python -m pip install pyinstaller"
}

if ($Clean) {
    if (Test-Path $BuildRoot) {
        Remove-Item $BuildRoot -Recurse -Force
    }
    if (Test-Path $DistRoot) {
        Remove-Item $DistRoot -Recurse -Force
    }
}

New-Item -ItemType Directory -Force -Path $BuildRoot | Out-Null
New-Item -ItemType Directory -Force -Path $DistRoot | Out-Null

python -m PyInstaller `
    --noconfirm `
    --clean `
    --distpath $DistRoot `
    --workpath $BuildRoot `
    $SpecPath

python packaging\verify_onefolder_dist.py --dist-root (Join-Path $DistRoot "CapacityOptimizer")

Write-Host ""
Write-Host "Build completed successfully."
Write-Host "Packaged app: $(Join-Path $DistRoot 'CapacityOptimizer')"

Pop-Location
