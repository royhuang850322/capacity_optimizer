# Capacity Optimizer User Guide

This guide is for business users who run the tool from the desktop launcher.

## 1. What You Need

- Windows computer
- Excel desktop application
- Valid `license.json` from support
- Tool package folder from GitHub release or internal delivery package

## 2. First-Time Setup (Customer Computer)

1. Copy the full tool folder to local disk (example: `D:\capacity_optimizer`).
2. If you are using source-mode delivery, run `runtime\setup_requirements.bat`.
3. Open `CapacityOptimizerLauncher.pyw` (or packaged `CapacityOptimizer.exe`).
4. Click `Initialize Workspace`.
5. Confirm workspace is created (default: `%LOCALAPPDATA%\CapacityOptimizer\` in packaged mode).

## 3. License Setup

Trial / unbound license:

1. Receive `license.json` from support.
2. Put it in `licenses\active\license.json` under workspace.

Machine-locked license:

1. Click `Generate Machine Fingerprint` in launcher.
2. Send generated file from `licenses\requests\` to support.
3. Receive signed `license.json`.
4. Put it in `licenses\active\license.json`.

## 4. Configure Run Settings in Launcher

Fill these key fields directly in launcher:

- `Project Root Folder`
- `Input Load Folder`
- `Input Master Folder`
- `Output Folder`
- `Output File Name`
- `Run Mode`
- `Start Year`, `Start Month`, `Horizon Months`
- `Direct Mode`, `Verbose`, `Skip Validation Errors`

Then click `Save Settings`.

## 5. Run and Review Results

1. Click `Run Optimizer`.
2. Open `Output Folder`.
3. Review generated Excel report files:
   - ModeA or ModeB workbook
   - comparison workbook when `Run_Mode = Both`
4. If run fails, open `Log Folder` and share latest run log with support.

## 6. Workspace Folder Map

```text
%LOCALAPPDATA%\CapacityOptimizer\
  Tooling Control Panel\
  docs\
  output\
  logs\
  licenses\
    active\
    requests\
```

## 7. Common Issues

- `Python was not found`:
  - install Python first (for source-mode workflow)
  - re-run `runtime\setup_requirements.bat`
- `License validation failed`:
  - check `licenses\active\license.json`
  - check expiry date and machine binding
- `Could not read control workbook`:
  - close workbook in Excel and retry
- `Failed to write output workbook`:
  - close output file in Excel
  - verify output folder write permission

## 8. Upgrade Guidance

- Replace app package with newer version.
- Keep workspace data unchanged.
- Do not delete old `licenses`, `output`, or workbook unless needed.

## 9. Support Checklist Before Raising Ticket

- Screenshot of error message
- Latest log file from `logs\`
- Launcher settings (or `launcher_settings.json` in workspace)
- Input files (`planner*_load`, `master_capacity`, `master_routing` if used)
