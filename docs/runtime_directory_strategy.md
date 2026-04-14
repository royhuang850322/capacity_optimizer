# Runtime Directory Strategy

This document captures the Milestone 2 runtime path rules for the Excel-first
Capacity Optimizer.

## Goal

Keep source-mode behavior stable while introducing a packaging-friendly path
model that also works under PyInstaller one-folder builds.

## Central Path Module

Runtime path resolution now lives in:

- `app/runtime_paths.py`

This module is responsible for resolving:

- application install directory
- bundled resources directory
- user workspace directory
- templates directory
- output directory
- logs directory
- license directory

## Source Mode

When the tool runs from source:

- install directory = repository root
- bundled resources directory = repository root
- user workspace directory = repository root

This preserves the current Excel-first development workflow and avoids
unnecessary behavior changes during the migration.

## Frozen / Packaged Mode

When the tool runs as a packaged executable:

- install directory = folder containing the executable
- bundled resources directory = `resources/`, `_internal/resources/`, or the
  install directory fallback
- user workspace directory =
  - `%LOCALAPPDATA%\\CapacityOptimizer`, or
  - `%USERPROFILE%\\.CapacityOptimizer` if `LOCALAPPDATA` is unavailable

The packaged application must treat bundled resources as read-only and user
workspace content as writable.

## Workspace Structure

Within the user workspace, the following structure is expected:

```text
<workspace>\
  Tooling Control Panel\
  docs\
  output\
  logs\
  licenses\
    active\
    requests\
  workspace_manifest.json
```

### Directory meanings

- `Tooling Control Panel\`
  - editable customer workbook copy
- `docs\`
  - customer-facing copies of quick-start, deployment, launcher, and Python install notes
- `output\`
  - generated Excel result workbooks
- `logs\`
  - runtime and support logs (CLI + launcher)
- `licenses\active\`
  - current active `license.json`
- `licenses\requests\`
  - generated machine fingerprint requests
- `workspace_manifest.json`
  - lightweight record of install location, workspace location, and initialization timestamps

## License Lookup Policy (Packaged Mode)

In packaged mode, the runtime validates licenses with a workspace-first search
order:

1. `<workspace>\licenses\active\license.json` (preferred)
2. `<control workbook Project_Root_Folder>\licenses\active\license.json`
3. `<install_dir>\licenses\active\license.json` (fallback only)
4. Legacy root-level `license.json` under each candidate root

## Logging Policy (Milestone 7)

- logging is centralized in `app/run_logging.py`
- run logs default to:
  - `<workspace>\logs\optimizer_run_YYYYMMDD_HHMMSS.log`
- launcher mode passes the same log path to CLI so both outputs stay in one file
- log files keep debug details for support:
  - path resolution
  - mode-level metrics
  - traceback details for unexpected failures
- user-facing terminal errors remain concise and include:
  - error code
  - suggested actions
  - log file path

This keeps machine-locked validation stable even when workbook project-root
settings are temporarily incorrect, while preserving backward compatibility
for older layouts.

## Resource Directories

Bundled read-only resources are expected to come from:

- `Data_Input\`
- workbook templates or workbook generation logic
- documentation copied into delivery packages as needed

In source mode these still resolve from the repository. In packaged mode they
must resolve from bundled application resources rather than the current working
directory.

## Milestone 5 Separation Rule

From this milestone onward, packaged runs should treat the install directory as
read-only application content and the workspace as the only writable home for:

- the editable control workbook
- copied sample/demo input data
- generated outputs
- logs
- active licenses and fingerprint requests
- copied customer-facing docs

Workspace initialization only creates missing files and folders. Existing user
files are preserved, which makes upgrades safer because new installs do not
overwrite the customer's workbook, data, or results.

## Current Compatibility Policy

Milestone 2 keeps these compatibility principles:

- source mode remains repository-root based
- packaged mode is ready for a separated user workspace
- legacy root-level `license.json` remains supported by the validator
- existing workbook-driven path settings remain intact

## Why This Matters

These rules reduce packaging risk by eliminating assumptions that:

- the current working directory is the repository root
- install directories are writable
- templates, licenses, outputs, and logs should all live together

This path model is the foundation for:

- launcher work in Milestone 3
- PyInstaller one-folder packaging in Milestone 4
- install-dir and workspace separation in Milestone 5
