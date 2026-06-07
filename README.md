# Capacity Optimizer

Capacity Optimizer is a Windows desktop capacity-planning tool.

Current workflow:

- Desktop GUI: `CapacityOptimizer.exe` in packaged mode, or `CapacityOptimizerLauncher.pyw` in source mode
- Input data: CSV / Excel files under `Data_Input`
- Optimization engine: Python + OR-Tools
- Output reports: Excel workbooks under `output`
- License: `licenses\active\license.json`

The old Excel control workbook UI has been retired. The historical workbook is archived under:

```text
Archive\legacy_excel_control_panel\Capacity_Optimizer_Control.xlsx
```

It is kept only for legacy reference and compatibility testing. Business users should not open it to run the tool.

## Quick Start

### Source Mode

```powershell
runtime\setup_requirements.bat
python CapacityOptimizerLauncher.pyw
```

In the launcher:

1. Click `Initialize Workspace`.
2. Put `license.json` under `licenses\active\license.json`.
3. Set run parameters directly in the launcher.
4. Click `Save Settings`.
5. Click `Run Optimizer`.
6. Open `Output Folder` or `Log Folder` from the launcher.

### Packaged Mode

Open:

```text
dist\CapacityOptimizer\CapacityOptimizer.exe
```

The launcher creates and uses a user workspace, typically:

```text
%LOCALAPPDATA%\CapacityOptimizer
```

## Main Files

```text
CapacityOptimizerLauncher.pyw
ProductAnalysisLauncher.pyw
WorkCenterAnalysisLauncher.pyw
Data_Input\
docs\
runtime\
packaging\
app\
tests\
Archive\legacy_excel_control_panel\
```

## Release / Maintenance

Set up development dependencies:

```powershell
.\scripts\bootstrap_dev.ps1
```

Run preflight:

```powershell
.\scripts\release_preflight.ps1
```

Run tests:

```powershell
python -m pytest
```

Build the main executable:

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Target CapacityOptimizer -Clean -CreateZip
```

## Documentation

- `docs\user_guide.md`
- `docs\desktop_launcher_usage.md`
- `docs\developer_guide.md`
- `docs\pyinstaller_onefolder_build.md`
- `docs\Capacity_Optimizer_v2.2.1_User_Guide_CN.docx`
