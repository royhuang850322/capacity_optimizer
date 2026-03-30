# Changelog

All notable changes to this project will be documented in this file.

The format is based on a simple release log:
- `Added` for new features
- `Changed` for behavior or workflow updates
- `Fixed` for bug fixes

## [v1.1.0] - 2026-03-30

Licensing and internal administration release for the Excel-first capacity optimizer.

### Added
- Offline license validation with signed `license.json`, expiry-date enforcement, and optional machine-locked authorization
- Machine fingerprint collection flow for customer computers
- Internal GUI license generator for issuing `trial / unbound` and `machine_locked` licenses without typing long PowerShell commands
- Managed internal license repository structure under `D:\RSCP_License_Admin\<CustomerName>\capacity_optimizer\...`
- License status fields written into result workbook `Run_Info`

### Changed
- Internal license tooling moved under `license_admin/`
- Project operations documents moved under `docs/`
- Readme and internal SOPs now point to the reorganized paths
- Tool version string updated to `v1.1.0`

### Fixed
- Dependency bootstrap and run scripts now align with the current Excel-first + license-controlled workflow
- Internal license generation defaults now point to the managed private-key and admin-repository paths

## [v1.0.1] - 2026-03-29

Planner-traceability release for the Excel-first capacity optimizer.

### Added
- Planner-level traceability in result outputs by splitting product-month results back to planner demand shares
- `Planner_Result_Summary` sheet in each result workbook
- `Planner_Product_Month` sheet in each result workbook
- `Planner_Compare` sheet in the standalone ModeA vs ModeB comparison workbook

### Changed
- `Allocation_Detail` now includes `PlannerName`
- `Allocation_Summary`, `Outsource_Summary`, and `Unmet_Summary` now include `PlannerName`
- Comparison workbook now includes planner-level side-by-side analysis
- Tool version string updated to `v1.0.1`

### Fixed
- Preserved total tons and KPI consistency after planner-level traceability split
- Added regression coverage for planner traceability outputs and summary workbook planner comparison

## [v1.0.0] - 2026-03-29

Initial managed release of the Excel-first capacity optimizer.

### Added
- Excel control workbook workflow centered on `Tooling Control Panel/Capacity_Optimizer_Control.xlsx`
- Excel report output for `ModeA`, `ModeB`, and comparison workbook generation when `Run_Mode = Both`
- Summary workbook `Summary of Mode A and Mode B_YYYYMMDD_HHMMSS.xlsx`
- Built-in sample data set under `Data_Input`
- Dependency bootstrap script `setup_requirements.bat`
- Runner script `run_optimizer.bat`
- Bilingual deployment guidance in the control workbook and IT checklist

### Changed
- Replaced the previous web-based interaction flow with an Excel-first workflow
- Standardized portable path handling with `Project_Root_Folder`, `Input_Load_Folder`, `Input_Master_Folder`, and `Output_Folder`
- Updated `requirements.txt` to require `openpyxl>=3.1.5`

### Fixed
- Strict validation for project root, input folders, planner files, and required master files before execution
- Dependency checks before optimizer execution to stop early when the local Python environment is incomplete
- Summary workbook naming to include timestamps and avoid overwriting previous runs
