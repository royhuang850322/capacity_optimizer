# Release v2.1.1

## Highlights

- `ModeBProductAnalysis` is now a lightweight companion tool that reads an existing `CapacityOptimizer` workspace
- Companion builds no longer initialize a duplicate `Data_Input / docs / licenses / output` tree beside the analysis exe
- The companion UI now lets the user choose the shared `CapacityOptimizer` working directory directly

## Assets

- `CapacityOptimizer-v2.1.1-win64.zip`
- `ModeBProductAnalysis-companion-v2.1.1-win64.zip`

## Validation

- `python -m pytest -p no:cacheprovider tests\test_modeb_customer_case_report.py tests\test_packaging_manifest.py tests\test_desktop_launcher.py -q`
- `python -m pytest -p no:cacheprovider tests\test_smoke_m8.py tests\test_regressions.py -q`
- `powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Target All -Clean -CreateZip -Version v2.1.1`
