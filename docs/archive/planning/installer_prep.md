# Installer Preparation (Milestone 9)

This document defines the installer-facing rules for shipping the packaged
Capacity Optimizer to enterprise Windows users.

## 1. Packaging Input

Installer input should come from the one-folder build output:

```text
dist\CapacityOptimizer\
  CapacityOptimizer.exe
  _internal\...
  _internal\resources\...
```

Do not re-structure runtime resources during install. Keep the packaged folder
layout intact.

## 2. Install Directory Policy

Preferred (no admin rights required):

```text
%LOCALAPPDATA%\Programs\CapacityOptimizer\
```

Fallback (if enterprise IT requires machine-wide install and admin rights):

```text
%ProgramFiles%\CapacityOptimizer\
```

Rules:
- Install directory is application-only (read-mostly).
- Do not store user workbook, outputs, logs, or active licenses in install dir.
- Runtime user data remains in workspace (see below), not in install path.

## 3. User Workspace Policy

Runtime workspace (already implemented by runtime path layer):

```text
%LOCALAPPDATA%\CapacityOptimizer\
  Tooling Control Panel\
  docs\
  output\
  logs\
  licenses\
    active\
    requests\
  workspace_manifest.json
```

Installer must not delete or overwrite this workspace during install/upgrade.

## 4. Shortcut Strategy

Create these shortcuts:

- Start Menu shortcut:
  - Name: `Chemical Capacity Optimizer`
  - Target: `<InstallDir>\CapacityOptimizer.exe`
- Optional Desktop shortcut (enterprise policy dependent):
  - Name: `Chemical Capacity Optimizer`
  - Target: `<InstallDir>\CapacityOptimizer.exe`

Recommended installer metadata:
- App display name: `Chemical Capacity Optimizer`
- Publisher: internal company publisher value
- Version: semantic version from release tag

No command-line arguments are required for normal users.

## 5. First-Run Initialization Behavior

On first launch, the app should initialize workspace automatically:

- create workspace directories if missing
- generate control workbook if missing
- copy sample data/docs only when missing
- preserve existing user files

Installer does not need to pre-populate workspace.
Initialization is app-driven and idempotent.

## 6. Upgrade Rules

Default upgrade mode: in-place application upgrade, preserve user data.

Rules:
- Replace install directory binaries/resources with new version.
- Keep `%LOCALAPPDATA%\CapacityOptimizer\` untouched.
- Preserve `licenses\active\license.json` in workspace.
- Preserve existing outputs and logs.
- Shortcut target updates automatically to new executable location.

Pre-upgrade checks:
- app process not running
- writable install target (or admin elevation if Program Files)

Post-upgrade checks:
- launcher starts
- workspace still points to existing user data
- run log is generated under workspace logs

## 7. Uninstall Rules

Default uninstall removes only installer-managed application files:

- remove install directory contents
- remove Start Menu/Desktop shortcuts

Default uninstall should **not** remove user workspace data.

Optional advanced uninstall checkbox:
- `Remove local workspace data (control workbook, outputs, logs, licenses)`
- default value: unchecked

## 8. Rollback Guidance

If a release fails in customer environment:

- reinstall previous version package
- keep workspace as-is
- verify with smoke test run from launcher

Because user data is workspace-based, rollback should not require data restore.

## 9. Installer Handoff Checklist

Before producing installer:

1. Build one-folder package.
2. Verify package layout (`packaging\verify_onefolder_dist.py`).
3. Validate launcher opens and workspace initializes.
4. Validate one minimal run writes output workbook and run log.
5. Confirm install/uninstall scripts do not delete workspace by default.

