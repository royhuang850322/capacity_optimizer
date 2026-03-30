# Chemical Capacity Optimizer

化工产能优化工具，当前版本采用：
- `CSV / Excel` 作为输入数据
- `Python + OR-Tools` 作为运算逻辑
- `Excel` 作为用户控制面板和结果报告载体

This project now uses an Excel-first workflow:
- CSV / Excel input files
- Python optimization logic
- Excel control workbook
- Excel result workbooks with report sheets

## 中文快速开始

### 1. 安装依赖

```bash
setup_requirements.bat
```

### 2. 生成机器指纹并申请授权

如果你准备发的是正式机绑版，新电脑第一次使用前先运行：

```bash
get_machine_fingerprint.bat
```

它会在项目根目录生成：

```text
machine_fingerprint.json
```

把这个文件发给 `RSCP`，拿到签名后的：

```text
license.json
```

然后把 `license.json` 放到项目根目录，也就是和 `main.py`、`run_optimizer.bat` 同级的位置。

如果你准备发的是短期试用版，也可以跳过机器指纹这一步，直接由 `RSCP` 签发一个：

- `trial`
- `unbound`

的 `license.json`，然后直接放到项目根目录。

### 3. 生成或刷新控制工作簿

```bash
python create_template.py
```

生成文件：

```text
Tooling Control Panel/Capacity_Optimizer_Control.xlsx
```

### 4. 在 Excel 里填写控制参数

打开 [Capacity_Optimizer_Control.xlsx](/C:/Users/super/capacity_optimizer/Tooling%20Control%20Panel/Capacity_Optimizer_Control.xlsx)，在 `Control_Panel` sheet 填写：

- `Project_Root_Folder`
- `Input_Load_Folder`
- `Input_Master_Folder`
- `Output_Folder`
- `Output_FileName`
- `Scenario_Name`
- `Start_Year`
- `Start_Month_Num`
- `Horizon_Months`
- `Run_Mode`
- `Direct_Mode`
- `Verbose`
- `Skip_Validation_Errors`

建议默认保持：

- `Project_Root_Folder = ..`
- `Input_Load_Folder = Data_Input`
- `Input_Master_Folder = Data_Input`
- `Output_Folder = output`

### 5. 运行工具

命令行方式：

```bash
python main.py --input-template "Tooling Control Panel/Capacity_Optimizer_Control.xlsx"
```

也可以直接双击：

[`run_optimizer.bat`](/C:/Users/super/capacity_optimizer/run_optimizer.bat)

它会按 `Tooling Control Panel/Capacity_Optimizer_Control.xlsx` 运行，并在成功后打开 `output/` 文件夹。

## English Quick Start

```bash
setup_requirements.bat
get_machine_fingerprint.bat
python create_template.py
python main.py --input-template "Tooling Control Panel/Capacity_Optimizer_Control.xlsx"
```

Before running, place a valid `license.json` in the project root and edit the `Control_Panel` sheet in `Tooling Control Panel/Capacity_Optimizer_Control.xlsx`.

## License Workflow

### Trial / Unbound

- RSCP generates a short-term unbound `license.json`
- Copy `license.json` into the project root
- Run the optimizer

### Machine-Locked

- Run `get_machine_fingerprint.bat` on the target computer
- Send `machine_fingerprint.json` to `RSCP`
- Receive the signed `license.json`
- Copy `license.json` into the project root
- Run the optimizer

The optimizer stops immediately when:

- `license.json` is missing
- the license signature is invalid
- the license has expired
- the machine fingerprint does not match the licensed computer

## Planning Modes

### ModeA

- Input: planner files + `master_capacity`
- No routing table is used
- Focus: internal allocation and residual unmet demand

### ModeB

- Input: planner files + `master_capacity` + routing master
- Routing file: `alternative_routing` or legacy `master_routing`
- Supports `Primary`, `Alternative`, and `Toller` logic
- Focus: internal allocation, outsourcing, and residual unmet demand

### Both

- Runs `ModeA` and `ModeB` in sequence
- Writes one Excel result workbook per mode
- Also writes a standalone comparison workbook with timestamp naming, for example: `Summary of Mode A and Mode B_20260327_144454.xlsx`

## Input Files

### Planner Files

Expected file names:

- `planner1_load`
- `planner2_load`
- `planner3_load`
- `planner4_load`
- `planner5_load`
- `planner6_load`

Supported extensions:

- `.csv`
- `.xlsx`
- `.xls`

Required columns:

- `Month`
- `PlannerName`
- `Product`
- `ProductFamily`
- `Plant`
- `Forecast_Tons`

Optional columns:

- `Scenario`
- `ResourceGroupOwner`
- `ScenarioVersion`
- `Comment`

### Capacity Master

Required columns:

- `Product`
- `WorkCenter`
- `Annual_Capacity_Tons`
- `Utilization_Target`

Optional columns:

- `Effective_From`
- `Effective_To`

### Routing Master

Used in `ModeB` only.

Columns:

- `Product` or `ProductFamily`
- `WorkCenter`
- `Priority`
- `EligibleFlag`
- `RouteType`
- `PenaltyWeight`

## Output Workbook

Each run writes a timestamped Excel workbook to the configured output folder.

When `Run_Mode = Both`, the tool also writes:

- `Summary of Mode A and Mode B_YYYYMMDD_HHMMSS.xlsx`

Report sheets:

- `Dashboard`
- `Monthly_Trend`
- `Bottleneck`
- `WC_Heatmap`
- `Product_Risk`

Raw / audit sheets:

- `Allocation_Detail`
- `Allocation_Summary`
- `Outsource_Summary`
- `Unmet_Summary`
- `WC_Load_Pct`
- `Binary_Feasibility`
- `Validation_Issues`
- `Run_Info`

Comparison workbook sheets:

- `Executive_Comparison`
- `Monthly_Trend_Compare`
- `Bottleneck_Compare`
- `WC_Heatmap_Compare`
- `Product_Risk_Compare`
- `Planner_Compare`
- `Run_Info`

## Control Workbook

The control workbook is the user interaction surface for this tool.

Main file:

- [Capacity_Optimizer_Control.xlsx](/C:/Users/super/capacity_optimizer/Tooling%20Control%20Panel/Capacity_Optimizer_Control.xlsx)

Generator:

- [create_template.py](/C:/Users/super/capacity_optimizer/create_template.py)

Runner:

- [main.py](/C:/Users/super/capacity_optimizer/main.py)

License helper:

- [get_machine_fingerprint.bat](/C:/Users/super/capacity_optimizer/get_machine_fingerprint.bat)

Internal signing helpers:

- [generate_license.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/generate_license.py)
- [generate_trial_license.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/generate_trial_license.py)
- [open_license_generator.bat](/C:/Users/super/capacity_optimizer/open_license_generator.bat)
- [license_generator_ui.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/license_generator_ui.py)
- [license_tools README](/C:/Users/super/capacity_optimizer/license_admin/license_tools/README.md)

License operation guides:

- [客户授权使用说明](/C:/Users/super/capacity_optimizer/docs/CUSTOMER_LICENSE_QUICKSTART_CN.md)
- [内部 License 发放 SOP](/C:/Users/super/capacity_optimizer/docs/INTERNAL_LICENSE_SOP_CN.md)

## Repository Structure

```text
.
|-- main.py
|-- data_loader.py
|-- optimizer.py
|-- validator.py
|-- output_writer.py
|-- result_analysis.py
|-- create_template.py
|-- create_sample_data.py
|-- get_machine_fingerprint.bat
|-- license_validator.py
|-- machine_fingerprint.py
|-- Data_Input/
|-- Tooling Control Panel/
|   `-- Capacity_Optimizer_Control.xlsx
|-- output/
|-- docs/
|   |-- CHANGELOG.md
|   |-- CUSTOMER_LICENSE_QUICKSTART_CN.md
|   |-- INTERNAL_LICENSE_SOP_CN.md
|   `-- IT_DEPLOYMENT_CHECKLIST_CN.md
|-- license_admin/
|   |-- license_tools/
|   `-- private_keys/
|-- tests/
|   `-- test_regressions.py
`-- run_optimizer.bat
```

## Sample Data

The repository includes demonstration input data under:

- [Data_Input](/C:/Users/super/capacity_optimizer/Data_Input)

You can regenerate the sample data with:

```bash
python create_sample_data.py
```

Data dictionary:

- [DATA_INPUT_GUIDE_CN.md](/C:/Users/super/capacity_optimizer/Data_Input/DATA_INPUT_GUIDE_CN.md)

## Tests

```bash
python -m unittest discover -s tests -v
```

## IT Deployment

For Windows workstation / shared-folder deployment, use:

- [IT_DEPLOYMENT_CHECKLIST_CN.md](/C:/Users/super/capacity_optimizer/docs/IT_DEPLOYMENT_CHECKLIST_CN.md)

## Changelog

- [CHANGELOG.md](/C:/Users/super/capacity_optimizer/docs/CHANGELOG.md)

## License

MIT. See [LICENSE](/C:/Users/super/capacity_optimizer/LICENSE).
