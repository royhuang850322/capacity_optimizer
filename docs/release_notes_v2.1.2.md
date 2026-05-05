# Release v2.1.2

## Highlights

- Archived historical handoff, planning, and old release-note documents under `docs/archive/`
- Archived the superseded standalone example product-report generator under `archive/legacy_code/`
- Promoted `TECHNICAL_REFERENCE_CN.md` into the active tracked documentation set
- Stopped tracking local runtime state files such as launcher settings and workspace manifest

## Active Documentation

- `README.md`
- `docs/CAPACITY_OPTIMIZER_USER_MANUAL_CN.md`
- `docs/TECHNICAL_REFERENCE_CN.md`
- `docs/developer_guide.md`
- `docs/user_guide.md`
- `docs/desktop_launcher_usage.md`
- `docs/runtime_directory_strategy.md`
- `docs/pyinstaller_onefolder_build.md`

## Validation

- `python -m pytest -p no:cacheprovider tests\test_modeb_customer_case_report.py tests\test_workcenter_analysis_report.py tests\test_packaging_manifest.py tests\test_desktop_launcher.py tests\test_regressions.py tests\test_runtime_paths.py tests\test_smoke_m8.py -q`
- `powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Target All -Clean -CreateZip -Version v2.1.2`
