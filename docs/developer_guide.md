# Capacity Optimizer Developer Guide

This guide is for internal engineers maintaining build, packaging, and release.

## 1. Repository Orientation

Key folders:

- `app/` core runtime, optimizer logic, launcher integration, report handling
- `tests/` unit, regression, packaging, and smoke tests
- `runtime/` helper batch scripts for source-mode operation
- `packaging/` PyInstaller spec and build scripts
- `docs/` deployment and process documentation
- `license_admin/` internal licensing and delivery tooling

## 2. Local Development Setup

Recommended:

- Python 3.11+ on Windows
- virtual environment per repository clone

Fast setup:

```powershell
.\scripts\bootstrap_dev.ps1
```

Manual setup:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

Dependency split:

- `requirements.txt` contains runtime packages needed by the application.
- `requirements-dev.txt` contains maintainer packages needed for tests, Word document generation, visual document rendering, and PyInstaller builds.

Document rendering checks require external tools in addition to Python packages:

- LibreOffice (`soffice`) for DOCX-to-PDF conversion.
- Poppler (`pdftoppm`) for PDF-to-image conversion.

Both tools must be discoverable on `PATH` for release document rendering.

## 3. Core Run Commands

Open the desktop launcher in source mode:

```powershell
python CapacityOptimizerLauncher.pyw
```

Developers can still exercise the legacy workbook CLI for regression coverage, but it is no longer the customer UI.

## 4. Desktop Entry

User-facing desktop entry:

- `CapacityOptimizerLauncher.pyw`

Responsibilities:

- workspace initialization
- launcher settings management
- trigger optimizer run with direct runtime config
- open output/log folders
- generate machine fingerprint request

## 5. Runtime Path Model

Central module:

- `app/runtime_paths.py`

Rules:

- source mode: workspace defaults to repository root
- packaged mode: workspace defaults to `%LOCALAPPDATA%\CapacityOptimizer`
- user-writable data lives in workspace only
- install directory is read-only application content

Related docs:

- `docs/runtime_directory_strategy.md`
- `docs/archive/planning/installer_prep.md`

## 6. Logging and Diagnostics

Central logging module:

- `app/run_logging.py`

Behavior:

- one run log per run under workspace `logs\`
- launcher + CLI write to same run log in launcher flow
- fatal errors include error code + suggested actions + log path

## 7. Testing

Run full suite:

```powershell
python -m pytest
```

Run smoke tests only (M8):

```powershell
python -m unittest tests.test_smoke_m8 -v
```

Smoke coverage includes:

- runtime paths/workspace initialization
- launcher settings and legacy workbook regression coverage
- minimal run + output workbook + log generation

## 8. Packaging (PyInstaller One-Folder)

Run the preflight check before building a release:

```powershell
.\scripts\release_preflight.ps1
```

Build:

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Clean
```

Verify dist layout:

```powershell
python packaging\verify_onefolder_dist.py --dist-root dist\CapacityOptimizer
```

Main references:

- `docs/pyinstaller_onefolder_build.md`
- `docs/archive/planning/installer_prep.md`

## 9. Installer Handoff Rules

- preserve one-folder packaged layout
- do not write user data into install directory
- preserve workspace data on upgrade/uninstall by default
- create Start Menu shortcut to `CapacityOptimizer.exe`

## 10. Release Checklist

1. update version/changelog
2. run `.\scripts\release_preflight.ps1`
3. build one-folder dist
4. verify dist layout
5. verify launcher flow on clean machine
6. publish tag/release notes

## 11. License and Security Notes

- Runtime validation uses embedded public key (`app/license_validator.py`).
- Private signing keys stay in internal admin environment only.
- Do not commit customer-specific `license.json` files.
