# Release v2.1.0

## Highlights

- Added the standalone `ModeBProductAnalysis` desktop tool for product-level report walkthroughs
- Localized more report content for Chinese workbook generation
- Added visible unmet-attribution detail sheets to the main report outputs
- Refreshed sample input data to better demonstrate routing, toller, and unmet behaviors
- Added dual-target PyInstaller packaging for both desktop launchers

## Assets

- `CapacityOptimizer-v2.1.0-win64.zip`
- `ModeBProductAnalysis-v2.1.0-win64.zip`

## Validation

- `python -m pytest -p no:cacheprovider tests\test_smoke_m8.py tests\test_desktop_launcher.py tests\test_regressions.py tests\test_modeb_customer_case_report.py tests\test_packaging_manifest.py -q`
- `powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1 -Target All -Clean -CreateZip -Version v2.1.0`
