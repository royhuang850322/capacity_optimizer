# PyInstaller One-Folder Build

This document captures the Milestone 4 packaging setup for building the
Capacity Optimizer as a Windows one-folder distribution.

## Build Target

The primary packaged output is a one-folder desktop distribution:

```text
dist\
  CapacityOptimizer\
    CapacityOptimizer.exe
    _internal\
      ...
      resources\
        Data_Input\
        docs\
```

PyInstaller commonly places collected data files under `_internal\resources\`
in one-folder mode. The runtime path layer also supports a top-level
`resources\` directory so that future installer layouts can override bundled
assets without changing application code.

The one-folder approach is preferred because it is more stable for:

- OR-Tools native dependencies
- cryptography
- openpyxl / pandas dependency trees
- bundled read-only resources

## Main Entry

The packaged application entry is:

- `CapacityOptimizer.exe`

This executable is built from:

- `CapacityOptimizerLauncher.pyw`

The launcher remains lightweight and is responsible for:

- initializing the user workspace
- creating or opening the control workbook
- generating machine fingerprint requests
- running the optimizer
- opening output and log folders

## Bundled Resources

The one-folder build collects these read-only resources into the packaged
resources directory, typically `_internal\resources\`:

- `Data_Input\`
- customer-facing docs from `docs\`
- `README.md`
- `LICENSE`

### Excel template note

The control workbook is not packaged as a fixed static template file. Instead,
it is generated on first run by:

- `app/create_template.py`

This avoids shipping a user-specific workbook and keeps the packaged template
logic aligned with the current Python code.

## Build Files

The packaging setup introduced in this milestone is:

- `packaging/CapacityOptimizer.spec`
- `packaging/build_onefolder.ps1`
- `packaging/verify_onefolder_dist.py`
- `build_support/packaging_manifest.py`

## Build Command

From the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1
```

If you want a clean rebuild:

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Clean
```

## Dependency Handling

The spec explicitly handles packaging-sensitive dependencies:

- hidden imports:
  - `ortools`
- dynamic libraries:
  - `ortools`
- copied package metadata:
  - `ortools`
  - `pandas`
  - `openpyxl`
  - `click`
  - `colorama`
  - `cryptography`

## Minimal Packaged Validation

After building:

1. Confirm the directory exists:
   - `dist\CapacityOptimizer\`
2. Run the layout verifier:

```powershell
python packaging\verify_onefolder_dist.py --dist-root dist\CapacityOptimizer
```

3. Launch:
   - `dist\CapacityOptimizer\CapacityOptimizer.exe`
4. In the launcher:
   - initialize the workspace
   - generate a machine fingerprint request
   - open the control workbook
   - run the optimizer after a valid license is present

## Risks and Current Limits

- This milestone prepares a stable build configuration, but the exact final
  packaged behavior still depends on PyInstaller being available in the build
  environment.
- The packaged launcher assumes Excel desktop is available on the customer
  machine for workbook interaction.
- No license file is bundled into the package; licenses remain user/workspace
  content.
- The user workspace is still validated in later milestones as we continue
  refining install-dir versus workspace behavior.

## Why This Is a Low-Risk Build Strategy

- one-folder instead of one-file
- launcher entry instead of BAT scripts
- read-only resources bundled separately under the packaged `resources\` root
- user data stays outside the packaged application folder
