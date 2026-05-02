# Changelog

All notable changes to this project will be documented in this file.

The format is based on a simple release log:
- `Added` for new features
- `Changed` for behavior or workflow updates
- `Fixed` for bug fixes

## [v2.1.1] - 2026-05-02

Shared-workspace companion packaging release for the ModeB product analysis tool.

### Changed
- `ModeBProductAnalysis` now reuses an existing `CapacityOptimizer` workspace instead of initializing its own full `Data_Input/docs/licenses/output` tree
- Companion UI now lets the user choose the shared `CapacityOptimizer` working directory and reads reports from that workspace
- Companion packaging now omits bundled customer workspace resources and ships as a lightweight companion package
- Tool version string updated to `v2.1.1`

### Fixed
- Companion builds now derive fallback `Data_Input` paths from the selected ModeB report location when the workbook's original `Run_Info` paths are unavailable
- PyInstaller verification now supports lightweight desktop companions that do not bundle report resources

## [v2.1.0] - 2026-05-02

ModeB product analysis release with refreshed sample data, localized reporting, and dual desktop packaging.

### Added
- Standalone `ModeBProductAnalysisLauncher.pyw` desktop tool for generating product-level analysis workbooks from existing `ModeB` outputs
- Visible `Unmet_Attribution_Detail` / `未满足回挂明细` sheet across generated reports
- Dual-target PyInstaller one-folder build flow for both `CapacityOptimizer` and `ModeBProductAnalysis`

### Changed
- Sample `Data_Input` datasets now better illustrate partial internal capacity, routing reroute, toller, and unmet behaviors
- Chinese report generation now localizes more visible report strings and follows the selected report language
- Packaged Windows executables now carry release version metadata and dedicated icons
- Tool version string updated to `v2.1.0`

### Fixed
- Excel report header merges now align with underlying table widths
- Report tables now consistently expose filterable Excel headers
- `ModeB` product analysis exports no longer trigger Excel content-recovery warnings on open

## [v1.1.3] - 2026-03-31

Dashboard filtering, pressure-load reporting, and control-workbook license visibility release.

### Added
- WorkCenter-filtered dashboard controls for `ModeA`, `ModeB`, and summary comparison workbooks
- A visible `License` sheet in `Capacity_Optimizer_Control.xlsx`
- Report-side pressure-load helpers for attributing `Unmet` and `Outsourced` demand by WorkCenter

### Changed
- Heatmap and bottleneck percentages now display against nameplate monthly capacity instead of utilization-limited capacity
- `ModeA` dashboard and heatmap unmet attribution now follows planner-resource ownership rules
- `ModeB` dashboard and heatmap attribution now sends outsourced tons to `Toller` and applies routing/capacity fallback rules for unmet demand
- Control workbook generation and refresh now show the currently detected license details
- Tool version string updated to `v1.1.3`

### Fixed
- Added validation for planner/product multi-resource conflicts and duplicate product-level routing definitions that break reporting attribution
- Delivery package export now refreshes the control workbook license page when a license file is bundled

## [v1.1.2] - 2026-03-30

Documentation and packaging polish release for the Excel-first capacity optimizer.

### Added
- Formatted Word user manual generator for rebuilding the customer/internal handbook from Markdown

### Changed
- Published the formatted Word user manual in the repository root
- Tool version string updated to `v1.1.2`

### Fixed
- Improved Word manual formatting so heading hierarchy, lists, and code blocks render more cleanly

## [v1.1.1] - 2026-03-30

Customer-delivery tooling and documentation release for the Excel-first capacity optimizer.

### Added
- Internal GUI delivery-package exporter with a one-click batch launcher
- Clean customer package export flow that copies only runtime files and can optionally include a signed license
- Root-level Word user manual covering customer-side and developer-side operations

### Changed
- Delivery package export is now available from both CLI and GUI workflows
- Delivery-package SOP and internal license SOP now point to the GUI exporter entry
- Tool version string updated to `v1.1.1`

### Fixed
- Reduced manual command-line steps for internal delivery preparation
- Added regression coverage for the delivery exporter GUI path setup

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
