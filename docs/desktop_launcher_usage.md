# Desktop Launcher Usage

The desktop launcher is now the primary customer workflow. Users no longer
need to run `python -m ...`, BAT files, or the Excel control workbook to run.

## Entry Point

The new desktop entry is:

- `CapacityOptimizerLauncher.pyw`

In source mode, double-clicking this file opens a small Windows launcher
window. In later packaging milestones, this same launcher logic is intended to
become the customer-facing executable.

## Workspace Model

The launcher works against the user workspace, not the packaged install folder.
In packaged mode the workspace is typically:

- `%LOCALAPPDATA%\CapacityOptimizer`

The install directory stays read-only while the launcher prepares user-editable
files inside the workspace.

## What the Launcher Does

The launcher supports the following customer actions:

- initialize the user workspace
- initialize sample data into the workspace when needed
- copy customer-facing support docs into the workspace when needed
- generate a machine fingerprint request file for machine-locked licensing
- configure all run parameters directly in launcher settings
- save launcher settings into workspace
- run the optimizer
- open the output folder
- open the log folder
- open the workspace folder

## Main Workflow

1. Open `CapacityOptimizerLauncher.pyw`
2. Click **Initialize Workspace** if needed
3. If you need a machine-locked license, click **Generate Machine Fingerprint**
4. Fill in folders and run settings directly in launcher
5. Click **Save Settings**
6. Click **Run Optimizer**
7. Open **Output Folder** or **Latest Log** as needed

## How the Launcher Connects to the Existing Flow

The launcher keeps the optimizer business flow intact:

- it prepares the workspace using `app/workspace_init.py`
- it writes machine fingerprint request files into `licenses\requests\`
- it builds runtime config directly from launcher fields
- it triggers shared optimization flow in `app/main.py` via `run_with_config(...)`
- it writes launcher output and CLI debug logs to a single run log file under `logs/`
- it keeps customer-editable content in the workspace instead of the install directory

## Log + Error Behavior (Milestone 7)

- every run creates a file log in `logs\optimizer_run_YYYYMMDD_HHMMSS.log`
- logs include debug details for path resolution, mode metrics, and failure tracebacks
- user-facing errors in CLI include:
  - short error code (for support)
  - concise summary and suggested actions
  - log file location for troubleshooting
- launcher users still see simple message boxes, while support can inspect the same log file

This means the customer experience changes, but the optimizer core and workbook
workflow remain the same.

## Why This Is a Low-Risk Step

- PySide6 Qt Widgets UI, backend logic unchanged
- existing CLI remains available for developers
- launcher does not change optimization rules
- launcher uses direct runtime config and shared optimizer pipeline
- logs are written to a predictable user-writable folder
