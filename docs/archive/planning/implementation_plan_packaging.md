# Packaging Implementation Plan

## Scope

This document is the Milestone 1 audit and packaging design for converting the current Excel-first Python tool into a Windows desktop deliverable that:

- does not require the customer to manage Python manually
- minimizes direct BAT/script usage for end users
- preserves the current Excel-first workflow
- remains suitable for enterprise Windows deployment
- uses PyInstaller `one-folder` as the primary packaging target

This milestone is analysis-only. No business logic changes are included here.

## Current Runtime Flow

### Customer-facing runtime chain

Current customer usage is driven by BAT scripts plus the Excel control workbook:

1. `runtime\check_python_setup.bat`
   - verifies that Python is available
   - attempts basic PATH repair if Python is installed but not visible
2. `runtime\setup_requirements.bat`
   - installs required Python packages with `pip`
   - creates `licenses\active` and `licenses\requests`
3. `runtime\get_machine_fingerprint.bat`
   - runs `python -m app.machine_fingerprint`
   - writes a machine fingerprint request file under `licenses\requests`
4. `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
   - user edits runtime configuration and sees the visible `License` sheet
5. `runtime\run_optimizer.bat`
   - validates Python and imports
   - checks for `licenses\active\license.json` or legacy root `license.json`
   - launches `python -m app.main --input-template "...Capacity_Optimizer_Control.xlsx"`
6. `app.main`
   - loads workbook config
   - validates the license
   - refreshes the workbook `License` sheet when possible
   - loads data
   - runs ModeA / ModeB optimization
   - writes Excel output workbooks

### Developer / internal runtime chain

Internal operations additionally use:

- `license_admin\open_license_generator.bat`
- `license_admin\license_tools\license_generator_ui.py`
- `license_admin\export_customer_package.py`
- `license_admin\open_delivery_exporter.bat`

These are not part of the intended customer runtime experience.

## Current Directory Structure and Key Resources

### Top-level directories

- `app/`
  - core Python application modules
- `runtime/`
  - customer-facing BAT entry points
- `Tooling Control Panel/`
  - Excel control workbook
- `Data_Input/`
  - sample/demo input data
- `output/`
  - generated Excel output workbooks
- `licenses/`
  - runtime license locations
- `docs/`
  - user and operational documentation
- `license_admin/`
  - internal license tooling and delivery export tooling
- `tests/`
  - regression and delivery package tests

### Critical runtime resources

- `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
  - control workbook
  - visible customer interaction surface
- `Data_Input\`
  - demo/sample data
- `licenses\active\license.json`
  - primary runtime license location
- `licenses\requests\`
  - machine fingerprint request output folder

### Core Python modules relevant to packaging

- `app\main.py`
  - main runtime entry point
- `app\create_template.py`
  - creates and refreshes the control workbook
- `app\data_loader.py`
  - configuration loading and data loading
- `app\optimizer.py`
  - OR-Tools optimization logic
- `app\output_writer.py`
  - Excel result workbook generation
- `app\license_validator.py`
  - offline signed license validation
- `app\machine_fingerprint.py`
  - Windows machine fingerprint generation

## External Dependencies

Current Python dependencies from `requirements.txt`:

- `ortools>=9.7`
- `openpyxl>=3.1.5`
- `pandas>=2.0`
- `click>=8.1`
- `colorama>=0.4`
- `cryptography>=45.0`

### Non-Python runtime assumptions

- Windows operating system
- Excel desktop application for full workbook interaction and formula/chart refresh
- ability to read Windows registry key:
  - `HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid`
- writable user-space folders for results, temporary files, logs, and license request files

## Current Packaging-Relevant Entry Points

### Customer-facing

- `runtime\run_optimizer.bat`
- `runtime\setup_requirements.bat`
- `runtime\get_machine_fingerprint.bat`
- `runtime\check_python_setup.bat`
- `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`

### Internal-only

- `license_admin\open_license_generator.bat`
- `license_admin\open_delivery_exporter.bat`
- `license_admin\export_customer_package.py`
- `license_admin\license_tools\*.py`

## Key Path and Resource Risks

### 1. Current working directory dependence in BAT scripts

Current BAT scripts use:

```text
cd /d "%~dp0\.."
```

This is stable in source form, but not ideal for packaged desktop delivery because:

- end-user flow still depends on BAT scripts
- Python startup is coupled to the repository layout
- the working directory becomes part of the runtime contract

### 2. Source-layout assumptions in Python modules

Examples:

- `app\create_template.py`
  - uses `ROOT_DIR = os.path.dirname(os.path.dirname(__file__))`
  - default output path points directly into `Tooling Control Panel`
- `license_admin\export_customer_package.py`
  - uses `REPO_ROOT = Path(__file__).resolve().parents[1]`

These are appropriate for source layout, but will become fragile once packaged because:

- `__file__` resolution changes under PyInstaller
- bundled resources may live in the executable distribution tree rather than the source tree
- user-writable output should not be coupled to install location

### 3. Mixed install-dir and user-data assumptions

Current design mixes concerns between:

- source/install directory
- editable control workbook
- demo data
- outputs
- license files

Today these commonly live under the project root:

- `Tooling Control Panel\`
- `Data_Input\`
- `output\`
- `licenses\`

This is acceptable in source mode, but high-risk in packaged mode because:

- packaged application folders may be read-only
- corporate users may install under protected locations
- upgrades could overwrite user data if directories are not separated

### 4. Control workbook path assumptions

The control workbook assumes a relative structure such as:

- `Project_Root_Folder = ..`
- `Input_Load_Folder = Data_Input`
- `Input_Master_Folder = Data_Input`
- `Output_Folder = output`

This is convenient in source mode, but after packaging:

- the control workbook should likely live in a user workspace, not beside the executable
- templates and user-edited copies should be separated
- workbook paths should no longer implicitly depend on the repository layout

### 5. License location ambiguity

Current runtime supports:

- preferred: `licenses\active\license.json`
- legacy fallback: root `license.json`

This dual-path behavior is good for backward compatibility, but it becomes a deployment risk because:

- packaged installs need one recommended, stable user-writable location
- multiple search locations can confuse support and customers

### 6. No central runtime path abstraction

Path logic is currently spread across:

- `app\create_template.py`
- `app\data_loader.py`
- `app\main.py`
- `app\license_validator.py`
- `runtime\*.bat`
- `license_admin\export_customer_package.py`

This increases packaging risk because:

- different code paths may compute different roots
- install-dir and user-dir separation is not centrally enforced
- PyInstaller compatibility is harder when path logic is duplicated

### 7. No built-in structured log directory yet

Current scripts print to console and some failures are user-friendly, but there is no centralized persistent runtime log policy for:

- launcher failures
- path initialization
- license issues
- workbook refresh failures
- file lock / Excel-open conflicts

This is a major enterprise supportability gap.

### 8. Excel workbook file-lock behavior

`app.main` already tolerates failure to refresh the workbook `License` sheet when the file is open in Excel.

This is good defensive behavior, but it highlights a packaging/runtime concern:

- desktop launchers must clearly explain when a workbook refresh could not be saved
- log output should capture this condition for support

### 9. Internal tooling currently lives in the same repository tree

The repository contains:

- customer runtime assets
- internal license tooling
- internal delivery export tooling

This is manageable in source control, but packaged deliverables must ensure:

- customer runtime does not include internal admin tools unless explicitly intended
- packaging scripts distinguish customer bundle vs developer/internal bundle

### 10. Presence of unrelated directories

The repository currently contains additional directories such as:

- `study_timer_app/`

This is not part of the optimizer runtime chain and must be explicitly excluded from packaging considerations and final build configuration.

## Excel Template and User Work File Separation Recommendation

This separation is strongly recommended for packaging:

### Bundled template assets

Should live in the packaged install/bundled resources area:

- pristine control workbook template
- optional pristine demo/sample data

### User workspace copies

Should live in a user-writable workspace:

- customer-editable control workbook
- working input copies if provided to the customer
- generated outputs
- logs
- runtime license files
- machine fingerprint requests

### Why this separation matters

- upgrades should not overwrite user-modified workbooks
- Program Files or other protected install locations may be read-only
- support can instruct users to replace or refresh templates without touching outputs

## One-Folder Packaging Strategy

### Recommended packaging target

Use PyInstaller `one-folder` mode as the primary build output.

### Why one-folder is preferred here

- better compatibility with OR-Tools and larger native dependency trees
- easier troubleshooting of missing DLL/import issues
- simpler handling of bundled templates and static resources
- less startup overhead than one-file self-extraction
- easier enterprise inspection and security review

### Proposed packaged structure

High-level target concept:

```text
dist/
  CapacityOptimizer/
    CapacityOptimizer.exe
    _internal/...
    resources/
      templates/
      sample_data/
      docs/
```

Exact directory names may vary, but the key principle is:

- executable and bundled read-only assets stay together in install/distribution area
- user-modified state does not live there

### Packaging implications

PyInstaller configuration will need to handle:

- `ortools`
- `cryptography`
- `pandas`
- `openpyxl`
- workbook template files
- sample/demo data if included
- documentation assets if customer bundle should contain them

## Recommended User Directory Plan

### Install directory

Purpose:

- packaged application binaries
- bundled read-only templates/resources

Examples:

- unpacked one-folder distribution location chosen by IT or the user
- later installer target such as:
  - `%LocalAppData%\Programs\CapacityOptimizer`
  - or corporate-managed install directory

### User workspace directory

Purpose:

- user-owned editable/runtime state

Recommended default base:

- `%USERPROFILE%\Documents\CapacityOptimizer`
  or
- `%LocalAppData%\CapacityOptimizer`

Recommended subdirectories:

```text
<workspace>\
  Tooling Control Panel\
  Data_Input\
  output\
  logs\
  licenses\
    active\
    requests\
```

### Recommendation

For enterprise usability, prefer:

- user workspace under a user-writable directory
- first-run initialization that copies the control workbook template into that workspace

## Recommended License Directory Plan

### Runtime/customer-facing

Recommended packaged/runtime license structure:

```text
<workspace>\licenses\
  active\
    license.json
  requests\
    machine_fingerprint_<Machine>.json
```

### Internal/admin-facing

Keep separate from packaged runtime:

- `license_admin/`
- signing keys
- internal issuance history

These should not be part of the customer runtime package.

## Recommended Log Directory Plan

Introduce a dedicated user-writable log directory:

```text
<workspace>\logs\
```

Suggested log categories:

- launcher log
- main run log
- license validation log
- initialization/path log
- workbook refresh/logical warning log

This is required for enterprise supportability and should be addressed in a later milestone.

## Windows Packaging and Enterprise Deployment Risks

### Risk: customer machines without Python

Current runtime depends on Python being installed, configured, and visible on PATH.

Mitigation:

- packaged executable must remove this requirement
- BAT-based Python bootstrap should no longer be part of the default customer flow

### Risk: Microsoft Store Python alias confusion

Current support scripts already show this problem in the field.

Mitigation:

- packaged executable eliminates Python PATH dependence
- fallback diagnostic docs remain useful for development mode only

### Risk: writing into install directory

Current source mode expects writable directories under the project root.

Mitigation:

- separate install directory from user workspace
- perform first-run initialization into a user-writable location

### Risk: PyInstaller resource lookup failures

Modules using `__file__` or repository-relative assumptions may fail to find templates/data after packaging.

Mitigation:

- create a centralized runtime path module
- explicitly support both source mode and frozen mode
- stop scattering install-root calculations

### Risk: hidden imports / native dependencies

`ortools` and `cryptography` may need explicit packaging attention.

Mitigation:

- add a reproducible PyInstaller spec/build script
- add smoke tests that run the packaged executable through a minimal workflow

### Risk: Excel workbook open/lock behavior

Refreshing or updating workbook sheets may fail if the workbook is open.

Mitigation:

- preserve non-fatal handling
- add log capture
- improve user-facing messaging in launcher/runtime

### Risk: internal tooling leakage into customer package

The repository contains internal tools that are not appropriate for the customer runtime bundle.

Mitigation:

- customer packaging script must explicitly whitelist customer runtime assets
- internal tooling stays out of final customer distribution

### Risk: upgrade overwriting customer work

If future packaged upgrades reuse the same directories without separation, customer data can be lost.

Mitigation:

- keep templates/install resources separate from workspace files
- do not overwrite user control workbooks and outputs during upgrades

## Deployment Architecture Summary

### Target desktop architecture

Recommended final architecture:

- packaged executable in install/distribution directory
- bundled read-only templates/resources in install/distribution directory
- user workspace under user-writable Windows location
- logs and license files under user workspace
- Excel used as the visible front-end workbook and result viewer

### Operational model

1. application starts from packaged EXE
2. runtime path layer resolves install/resources vs workspace
3. workspace is initialized if missing
4. control workbook is opened or copied into workspace
5. main optimization process runs using workspace paths
6. outputs and logs are written to workspace

## Subsequent Milestone Plan

### Milestone 2: Runtime path governance

Goal:

- add one centralized runtime path module
- support both source mode and frozen mode
- define install dir, bundled resources dir, workspace dir, outputs dir, logs dir, and license dir

### Milestone 3: Formal desktop launcher

Goal:

- add a customer-facing executable launcher
- support:
  - open/copy control workbook
  - run optimization
  - open output directory
  - open log directory

### Milestone 4: PyInstaller one-folder build

Goal:

- add PyInstaller spec/build script
- package hidden imports and resource files
- produce a reproducible one-folder build

### Milestone 5: Install dir / workspace separation

Goal:

- initialize a user workspace on first run
- ensure templates, outputs, logs, and license files live outside install dir

### Milestone 6: License compatibility in packaged mode

Goal:

- make license lookup, machine fingerprint generation, and packaged path behavior consistent

### Milestone 7: Logging and diagnostics

Goal:

- add persistent logs
- improve customer-facing failure messages
- improve support diagnostics

### Milestone 8: Smoke and regression coverage

Goal:

- add tests for path initialization, template copy, output write, and packaged-resource lookup

### Milestone 9: Installer preparation

Goal:

- prepare packaging/install assumptions for a future Windows installer

### Milestone 10: Customer and internal docs

Goal:

- finalize build, packaging, installation, runtime, upgrade, and troubleshooting documentation

## Suggested Commit Message

```text
docs: add packaging implementation plan for Windows desktop delivery
```
