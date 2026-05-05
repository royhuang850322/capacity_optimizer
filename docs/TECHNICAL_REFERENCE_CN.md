# Capacity Optimizer Technical Reference

本文档从程序员视角总结当前项目的技术选型、系统架构、GUI 方案、license 机制、打包交付方式，以及后续继续开发时最值得保留的经验。

适用场景：

- 后续维护和迭代本项目
- 新开发者快速理解当前系统
- 复盘本项目的工程方案并迁移到类似工具

## 1. Project Positioning

这是一个典型的：

- 本地运行
- Windows 桌面交付
- Excel-first
- Python 后台求解
- 离线 license 控制

的企业优化工具。

它不是 Web 系统，也不是数据库驱动平台。它的设计核心是：保留业务用户熟悉的 Excel 工作方式，同时把复杂的数据处理、优化求解、报告生成和授权控制放到 Python 后端里。

## 2. Technology Stack

### 2.1 Primary Language

- Python

### 2.2 Core Libraries

- `pandas`
  Used for CSV / Excel data loading, cleaning, joins, aggregation, and result shaping.
- `openpyxl`
  Used for Excel workbook read/write, styling, charts, conditional formatting, and report generation.
- `ortools`
  Used for optimization solving and is the core engine of the tool.
- `click`
  Used for CLI entry points.
- `PySide6`
  Used for the Windows desktop GUI launcher.
- `cryptography`
  Used for license signing and validation.

### 2.3 Delivery Form

- Source mode: Python environment runs repository code directly
- Customer mode: PyInstaller one-folder Windows EXE package

## 3. System Architecture

当前项目采用的是“模块化单体应用”架构。它没有拆成前后端服务，而是按职责分成多个清晰模块。

主要分层如下：

1. Input Layer
   CSV / Excel input files from planners and master data
2. Orchestration Layer
   Config loading, runtime path governance, workspace initialization, license checks, and run coordination
3. Optimization Layer
   OR-Tools solving for ModeA / ModeB
4. Analysis and Reporting Layer
   Result shaping, dashboard facts, heatmaps, risk summaries, and workbook output
5. Desktop Interaction Layer
   PySide6 launcher for customer-facing operation
6. Delivery Layer
   PyInstaller packaging and release artifacts

这种架构的优点是：

- GUI、求解、报表没有混在一起
- 开发态和打包态可以共用同一套核心逻辑
- 后续更换 UI 或扩展 solver 时不需要重写整个系统

## 4. Core Modules

### 4.1 Main Entry

- [app/main.py](/C:/Users/super/capacity_optimizer/app/main.py)

Responsibilities:

- CLI entry
- Config loading
- License validation
- Input validation
- Run orchestration for ModeA / ModeB and capacity basis
- Workbook output trigger

### 4.2 Desktop Launcher

- [app/desktop_launcher.py](/C:/Users/super/capacity_optimizer/app/desktop_launcher.py)

Responsibilities:

- Provide customer-facing Windows GUI
- Save launcher settings
- Manage a single workspace root
- Initialize workspace
- Run optimizer
- Open output / logs / license request folders

Key design lesson:

- GUI should orchestrate workflow
- GUI should not contain business solving logic

这是当前项目里非常值得继续保持的一条边界。

### 4.3 Data Loading

- [app/data_loader.py](/C:/Users/super/capacity_optimizer/app/data_loader.py)

Responsibilities:

- Read planner files
- Read master data
- Normalize column names
- Filter by scenario
- Support dual capacity basis loading:
  - `Max Capacity Ton`
  - `Planner Capacity Ton`
  - backward compatibility with legacy `Designed Capacity Ton`

Key lesson:

- 输入层要优先做兼容和标准化
- 不要让新需求直接破坏旧数据可运行性

### 4.4 Optimizer

- [app/optimizer.py](/C:/Users/super/capacity_optimizer/app/optimizer.py)

Responsibilities:

- Solve ModeA
- Solve ModeB
- Apply capacity constraints
- Apply routing constraints
- Generate allocation / outsourced / unmet outcomes

Key lesson:

- 对新业务需求，优先通过“多次运行 + 编排层扩展”来实现
- 不优先在 solver 内部堆叠过多分支逻辑

双 capacity basis 的实现就属于这一类工程决策。

### 4.5 Result Analysis

- [app/result_analysis.py](/C:/Users/super/capacity_optimizer/app/result_analysis.py)
- [app/load_pressure.py](/C:/Users/super/capacity_optimizer/app/load_pressure.py)

Responsibilities:

- Monthly aggregation
- Planner traceability
- Product risk analysis
- WorkCenter load analysis
- Dashboard facts
- Heatmap facts

### 4.6 Workbook Output

- [app/output_writer.py](/C:/Users/super/capacity_optimizer/app/output_writer.py)

Responsibilities:

- Write single mode workbooks
- Write ModeA / ModeB comparison workbook
- Write dual capacity comparison views
- Build sheets, formatting, charts, slicers or filter-friendly tables

这是项目里最贴近最终用户体验的一层，也是复杂度较高的一层。

### 4.7 Runtime Paths

- [app/runtime_paths.py](/C:/Users/super/capacity_optimizer/app/runtime_paths.py)

Responsibilities:

- Define install dir, bundled resources dir, workspace, logs, output, licenses, docs, and related paths in one place
- Support both source mode and packaged mode

Key lesson:

- 路径必须集中治理
- 不要把路径判断散落在多个业务文件里

### 4.8 Workspace Initialization

- [app/workspace_init.py](/C:/Users/super/capacity_optimizer/app/workspace_init.py)

Responsibilities:

- Create folder structure on first run
- Initialize `Data_Input`
- Initialize `docs`
- Write `workspace_manifest.json`

Current rule set:

- Launcher only keeps one editable workspace root
- `Initialize Workspace` creates all derived folders under that root
- Control workbook is no longer auto-created by workspace initialization

## 5. GUI Strategy

### 5.1 Why PySide6

This project went through multiple UI stages:

- Streamlit
- Tkinter
- finally rebuilt with PySide6

Reasons for choosing PySide6:

- Looks closer to a formal enterprise desktop application
- Better layout and component control than Tkinter
- More suitable than a browser-based UI for offline customer deployment
- Easier to package as a Windows EXE

### 5.2 GUI Responsibility Boundary

The GUI should do:

- Parameter input
- Path saving
- Workspace initialization
- Run triggering
- Open logs / output / license request folders

The GUI should not do:

- Direct optimization solving
- Direct report calculation
- Core business logic branching

这条边界后续必须继续保持。

### 5.3 GUI Evolution Lessons

本项目在 launcher 重构中得到的经验：

- 不要把所有设置堆到一个超大页面
- 企业用户更适合“分区 + 明确主动作”的桌面界面
- 路径配置要尽量收敛成一个主路径
- 不要让用户同时理解多个彼此关联的目录

当前已收敛到：

- 一个可编辑的主 workspace 路径
- `Data_Input` 和 `output` 作为派生路径显示
- `Save Settings` 用于确认设置
- `Initialize Workspace` 用于创建目录结构

这比让用户手工配置多个独立路径稳定得多。

## 6. Input and Output Design

### 6.1 Inputs

输入是文件驱动，而不是数据库驱动。

Typical inputs include:

- planner files
- `master_capacity.csv`
- `master_routing.csv`
- related routing or capacity master files

Advantages:

- Easy for customers to maintain
- Easy to deliver
- Easy to export and archive

Risks:

- Strong dependence on column names
- Strong dependence on file format consistency
- Requires clear pre-validation

### 6.2 Outputs

输出是 Excel workbooks，而不是网页。

Advantages:

- Familiar to customers
- Easy to forward and print
- Easy to filter, sort, and continue analysis in Excel

Current output themes include:

- dashboard
- monthly analysis
- product risk
- planner summary
- allocation summary
- workcenter heatmap
- ModeA / ModeB comparison summary

## 7. Optimization and Business Logic Strategy

### 7.1 Modes

The tool currently supports at least:

- ModeA
- ModeB

### 7.2 Dual Capacity Basis

The current business evolution introduced two capacity baselines:

- `Max Capacity Ton`
- `Planner Capacity Ton`

The important implementation approach here is:

- Run the model separately by basis
- Compare results in reporting
- Do not overload the core solver with unnecessary branching if orchestration can handle it cleanly

这是一条值得在后续复杂需求中继续复用的原则。

### 7.3 Reporting Philosophy

报表层的职责不只是“导出结果”，而是把：

- KPI
- risk
- bottleneck
- planner traceability
- heatmap
- comparison

全部用 Excel-first 的方式交付给最终用户。

这也是该项目区别于纯脚本工具的重要地方。

## 8. Runtime Path Governance

这是当前项目最值得复用的工程实践之一。

Current path model clearly distinguishes:

- app install dir
- bundled resources dir
- workspace dir
- `Data_Input`
- `output`
- `logs`
- `licenses`
- `docs`

Recent simplification:

- Launcher now exposes only one editable workspace root
- All operational folders are derived from that root

Key lesson:

- 路径规则要集中
- 路径配置要尽量少
- 用户写入内容必须进入用户可写 workspace
- 安装目录尽量视为只读应用内容

## 9. License Architecture

### 9.1 Model

The project uses an offline signed license model:

1. Customer generates a machine fingerprint
2. Internal admin tooling signs and issues a `license.json`
3. Customer places the active license in the expected folder
4. Runtime validates signature, validity period, product scope, and machine binding

### 9.2 Why This Works

This approach is suitable for enterprise environments because:

- It works offline
- It avoids a live license server dependency
- It supports trial and formal delivery
- It supports machine-bound and unbound issuance models

### 9.3 Key Lesson

License should be treated as a process, not just a file.

That includes:

- request generation
- signing
- issuing
- archiving
- replacing
- diagnosing failures

## 10. Packaging and Delivery

### 10.1 Packaging Choice

Current delivery uses:

- PyInstaller
- one-folder mode

Relevant files:

- [packaging/CapacityOptimizer.spec](/C:/Users/super/capacity_optimizer/packaging/CapacityOptimizer.spec)
- [packaging/build_onefolder.ps1](/C:/Users/super/capacity_optimizer/packaging/build_onefolder.ps1)
- [build_support/packaging_manifest.py](/C:/Users/super/capacity_optimizer/build_support/packaging_manifest.py)

### 10.2 Why One-Folder

One-folder is preferred because:

- Customers do not need to install Python
- Resource files are easier to ship and inspect
- It is more stable than one-file for this type of app
- Diagnostics are easier when packaging-related problems occur

### 10.3 Delivery Recommendation

Recommended delivery split:

- Development delivery: source code + Python environment
- Customer delivery: packaged one-folder bundle
- External handoff: zip package built around the one-folder distribution

## 11. Testing Strategy

This project is not a web service, but it still needs strong regression confidence.

Current useful test types include:

- regression tests
- smoke tests
- launcher tests
- packaging verification

The most important tests are not high-throughput or performance tests. They are:

- Can the app start
- Can the workspace initialize
- Can paths resolve correctly
- Can reports be written
- Can packaged mode find resources
- Can license validation still work

这些测试更贴近真实交付风险。

## 12. Key Engineering Lessons

1. Do not force Excel users into a web workflow when Excel-first is the real fit.
2. Keep GUI focused on orchestration, not business solving.
3. Centralize runtime path governance.
4. Extend orchestration before rewriting the solver.
5. Treat license as a complete operational process.
6. Prefer one-folder packaging for Windows enterprise delivery.
7. Aim tests at deployment risks, not only function-level correctness.

## 13. Suggested Direction for Future Development

If this project continues to evolve, the most stable path is:

- Keep Python + OR-Tools as the solving stack
- Keep Excel as the primary interaction and reporting artifact
- Keep PySide6 as the desktop launcher technology
- Continue strengthening runtime path governance
- Continue improving report usability inside Excel
- Keep offline signed license architecture
- Keep PyInstaller one-folder as the main delivery form

## 14. Related Documents

- [Developer Guide](/C:/Users/super/capacity_optimizer/docs/developer_guide.md)
- [User Guide](/C:/Users/super/capacity_optimizer/docs/user_guide.md)
- [Desktop Launcher Usage](/C:/Users/super/capacity_optimizer/docs/desktop_launcher_usage.md)
- [Runtime Directory Strategy](/C:/Users/super/capacity_optimizer/docs/runtime_directory_strategy.md)
- [Installer Prep (Archived)](/C:/Users/super/capacity_optimizer/docs/archive/planning/installer_prep.md)
