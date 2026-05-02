param(
    [ValidateSet("CapacityOptimizer", "ModeBProductAnalysis", "All")]
    [string]$Target = "CapacityOptimizer",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = (Resolve-Path (Join-Path $ScriptDir "..")).Path
$BuildRoot = Join-Path $ProjectRoot "build\pyinstaller"
$DistRoot = Join-Path $ProjectRoot "dist"

$Targets = @(
    @{
        Name = "CapacityOptimizer"
        SpecPath = Join-Path $ScriptDir "CapacityOptimizer.spec"
    },
    @{
        Name = "ModeBProductAnalysis"
        SpecPath = Join-Path $ScriptDir "ModeBProductAnalysis.spec"
    }
)

function Invoke-OneFolderBuild {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BuildTarget
    )

    $AppName = $BuildTarget.Name
    $SpecPath = $BuildTarget.SpecPath

    Write-Host ""
    Write-Host "Building $AppName"
    Write-Host "Spec file  : $SpecPath"

    python -m PyInstaller `
        --noconfirm `
        --clean `
        --distpath $DistRoot `
        --workpath $BuildRoot `
        $SpecPath

    python packaging\verify_onefolder_dist.py `
        --dist-root (Join-Path $DistRoot $AppName) `
        --app-name $AppName

    Write-Host "Packaged app: $(Join-Path $DistRoot $AppName)"
}

Write-Host "Capacity Optimizer - PyInstaller One-Folder Build"
Write-Host "Project root : $ProjectRoot"
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

try {
    if ($Target -eq "All") {
        foreach ($BuildTarget in $Targets) {
            Invoke-OneFolderBuild -BuildTarget $BuildTarget
        }
    } else {
        $SelectedTarget = $Targets | Where-Object { $_.Name -eq $Target } | Select-Object -First 1
        if (-not $SelectedTarget) {
            throw "Unknown build target: $Target"
        }
        Invoke-OneFolderBuild -BuildTarget $SelectedTarget
    }
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "Build completed successfully."
