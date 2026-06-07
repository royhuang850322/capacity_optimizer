[CmdletBinding()]
param(
    [string]$Python = "",
    [switch]$SkipTests,
    [switch]$SkipDocumentRenderChecks
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $RepoRoot

$Failures = New-Object System.Collections.Generic.List[string]
$Warnings = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    param([string]$Message)
    $script:Failures.Add($Message) | Out-Null
    Write-Host "[FAIL] $Message" -ForegroundColor Red
}

function Add-Warning {
    param([string]$Message)
    $script:Warnings.Add($Message) | Out-Null
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Add-Ok {
    param([string]$Message)
    Write-Host "[ OK ] $Message" -ForegroundColor Green
}

function Test-PythonModule {
    param(
        [string]$PythonExe,
        [string]$ModuleName,
        [string]$PackageName
    )

    & $PythonExe -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Add-Ok "Python module available: $ModuleName"
    } else {
        Add-Failure "Missing Python module '$ModuleName' from package '$PackageName'"
    }
}

function Test-Tool {
    param(
        [string[]]$Names,
        [string]$Purpose
    )

    foreach ($Name in $Names) {
        $Command = Get-Command -Name $Name -ErrorAction SilentlyContinue
        if ($Command) {
            Add-Ok "$Purpose tool available: $Name"
            return $true
        }
    }

    Add-Failure "$Purpose tool missing. Tried: $($Names -join ', ')"
    return $false
}

if ([string]::IsNullOrWhiteSpace($Python)) {
    $VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $VenvPython) {
        $PythonExe = $VenvPython
    } else {
        $PythonExe = "python"
    }
} else {
    $PythonExe = $Python
}

Write-Host "Capacity Optimizer release preflight"
Write-Host "Repository: $RepoRoot"
Write-Host "Python    : $PythonExe"
Write-Host ""

& $PythonExe --version
if ($LASTEXITCODE -ne 0) {
    Add-Failure "Python is not executable: $PythonExe"
} else {
    Add-Ok "Python executable works"
}

$RuntimeModules = @(
    @{ Module = "ortools"; Package = "ortools" },
    @{ Module = "openpyxl"; Package = "openpyxl" },
    @{ Module = "pandas"; Package = "pandas" },
    @{ Module = "click"; Package = "click" },
    @{ Module = "colorama"; Package = "colorama" },
    @{ Module = "cryptography"; Package = "cryptography" },
    @{ Module = "PySide6"; Package = "PySide6" }
)

$DevModules = @(
    @{ Module = "pytest"; Package = "pytest" },
    @{ Module = "PyInstaller"; Package = "pyinstaller" },
    @{ Module = "docx"; Package = "python-docx" },
    @{ Module = "pdf2image"; Package = "pdf2image" },
    @{ Module = "PIL"; Package = "pillow" },
    @{ Module = "win32com"; Package = "pywin32" }
)

Write-Host ""
Write-Host "Checking Python runtime dependencies..."
foreach ($Item in $RuntimeModules) {
    Test-PythonModule -PythonExe $PythonExe -ModuleName $Item.Module -PackageName $Item.Package
}

Write-Host ""
Write-Host "Checking Python development/release dependencies..."
foreach ($Item in $DevModules) {
    Test-PythonModule -PythonExe $PythonExe -ModuleName $Item.Module -PackageName $Item.Package
}

if (-not $SkipDocumentRenderChecks) {
    Write-Host ""
    Write-Host "Checking document rendering tools..."
    [void](Test-Tool -Names @("soffice", "libreoffice") -Purpose "DOCX to PDF")
    [void](Test-Tool -Names @("pdftoppm") -Purpose "PDF to image")
} else {
    Add-Warning "Document rendering tool checks were skipped."
}

if (-not $SkipTests) {
    Write-Host ""
    Write-Host "Running test suite..."
    & $PythonExe -m pytest
    if ($LASTEXITCODE -eq 0) {
        Add-Ok "pytest passed"
    } else {
        Add-Failure "pytest failed"
    }
} else {
    Add-Warning "Test suite was skipped."
}

Write-Host ""
if ($Failures.Count -gt 0) {
    Write-Host "Preflight failed with $($Failures.Count) issue(s)." -ForegroundColor Red
    Write-Host "Install Python packages with:"
    Write-Host "  .\scripts\bootstrap_dev.ps1"
    Write-Host "Install external document tools and ensure their bin directories are on PATH:"
    Write-Host "  LibreOffice: provides soffice"
    Write-Host "  Poppler: provides pdftoppm"
    exit 1
}

if ($Warnings.Count -gt 0) {
    Write-Host "Preflight passed with $($Warnings.Count) warning(s)." -ForegroundColor Yellow
} else {
    Write-Host "Preflight passed." -ForegroundColor Green
}
