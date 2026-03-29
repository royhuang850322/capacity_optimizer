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

### 2. 生成或刷新控制工作簿

```bash
python create_template.py
```

生成文件：

```text
Tooling Control Panel/Capacity_Optimizer_Control.xlsx
```

### 3. 在 Excel 里填写控制参数

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

### 4. 运行工具

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
python create_template.py
python main.py --input-template "Tooling Control Panel/Capacity_Optimizer_Control.xlsx"
```

Edit the `Control_Panel` sheet in `Tooling Control Panel/Capacity_Optimizer_Control.xlsx` before running.

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
- `Run_Info`

## Control Workbook

The control workbook is the user interaction surface for this tool.

Main file:

- [Capacity_Optimizer_Control.xlsx](/C:/Users/super/capacity_optimizer/Tooling%20Control%20Panel/Capacity_Optimizer_Control.xlsx)

Generator:

- [create_template.py](/C:/Users/super/capacity_optimizer/create_template.py)

Runner:

- [main.py](/C:/Users/super/capacity_optimizer/main.py)

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
|-- Data_Input/
|-- Tooling Control Panel/
|   `-- Capacity_Optimizer_Control.xlsx
|-- output/
|-- tests/
|   `-- test_regressions.py
|-- IT_DEPLOYMENT_CHECKLIST_CN.md
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

- [IT_DEPLOYMENT_CHECKLIST_CN.md](/C:/Users/super/capacity_optimizer/IT_DEPLOYMENT_CHECKLIST_CN.md)

## License

MIT. See [LICENSE](/C:/Users/super/capacity_optimizer/LICENSE).
