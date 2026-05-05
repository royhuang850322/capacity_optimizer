# 项目交接说明

本文件用于把当前 `capacity_optimizer` 项目交接给下一位开发者，尤其适用于：

- 另一位开发者继续维护本项目
- 另一位使用 Codex 的开发者继续推进本项目
- 需要把当前项目状态、运行方式、打包方式、license 机制、后续开发重点一次性交代清楚

## 1. 项目定位

当前项目是一个：

- 本地运行的 Windows 桌面工具
- Excel-first 的业务优化工具
- Python 负责后台编排、数据处理和 OR-Tools 求解
- Excel 负责结果承载和最终业务报告
- PySide6 负责 launcher GUI
- 支持离线 license 和 Windows EXE 打包交付

当前产品名称：

- 英文：`Capacity Optimizer`
- 中文：`产能优化工具`

## 2. 当前核心能力

截至当前版本，项目已经具备以下能力：

- 从 `Data_Input` 读取 planner 和主数据
- 支持 `ModeA`、`ModeB`、`Both`
- 支持 `Max Capacity Ton` 与 `Planner Capacity Ton` 双口径对比
- 输出：
  - `ModeA` 单独报告
  - `ModeB` 单独报告
  - `Summary of Mode A and Mode B` 对比报告
- 支持中英文 GUI
- 支持中英文报告输出
- 支持离线 license 校验
- 支持 PyInstaller one-folder 打包

## 3. 建议先读的文档

下一位开发者接手时，建议按这个顺序先阅读：

1. [README.md](/C:/Users/super/capacity_optimizer/README.md)
2. [TECHNICAL_REFERENCE_CN.md](/C:/Users/super/capacity_optimizer/docs/TECHNICAL_REFERENCE_CN.md)
3. [developer_guide.md](/C:/Users/super/capacity_optimizer/docs/developer_guide.md)
4. [runtime_directory_strategy.md](/C:/Users/super/capacity_optimizer/docs/runtime_directory_strategy.md)
5. [pyinstaller_onefolder_build.md](/C:/Users/super/capacity_optimizer/docs/pyinstaller_onefolder_build.md)
6. [NEXT_PROJECT_REUSE_GUIDE_CN.md](/C:/Users/super/capacity_optimizer/docs/NEXT_PROJECT_REUSE_GUIDE_CN.md)
7. [NEW_PROJECT_KICKOFF_PROMPT_CN.md](/C:/Users/super/capacity_optimizer/docs/NEW_PROJECT_KICKOFF_PROMPT_CN.md)

如果是第一次接手当前项目，前 5 份必须先读完。

## 4. 仓库中最重要的目录和文件

### 4.1 核心源码

- [app/main.py](/C:/Users/super/capacity_optimizer/app/main.py)
  主运行入口、模式编排、license 校验、报表输出触发
- [app/desktop_launcher.py](/C:/Users/super/capacity_optimizer/app/desktop_launcher.py)
  客户端桌面 GUI 启动器
- [app/output_writer.py](/C:/Users/super/capacity_optimizer/app/output_writer.py)
  Excel 报表输出核心
- [app/data_loader.py](/C:/Users/super/capacity_optimizer/app/data_loader.py)
  输入数据读取与标准化
- [app/optimizer.py](/C:/Users/super/capacity_optimizer/app/optimizer.py)
  OR-Tools 求解逻辑
- [app/result_analysis.py](/C:/Users/super/capacity_optimizer/app/result_analysis.py)
  结果汇总与分析
- [app/load_pressure.py](/C:/Users/super/capacity_optimizer/app/load_pressure.py)
  WorkCenter 负载、heatmap、dashboard fact 相关逻辑
- [app/runtime_paths.py](/C:/Users/super/capacity_optimizer/app/runtime_paths.py)
  运行时路径统一治理
- [app/workspace_init.py](/C:/Users/super/capacity_optimizer/app/workspace_init.py)
  工作区初始化
- [app/i18n.py](/C:/Users/super/capacity_optimizer/app/i18n.py)
  中英文文本映射

### 4.2 打包与交付

- [packaging](/C:/Users/super/capacity_optimizer/packaging)
- [build_support](/C:/Users/super/capacity_optimizer/build_support)
- [delivery_packages](/C:/Users/super/capacity_optimizer/delivery_packages)
- [dist](/C:/Users/super/capacity_optimizer/dist)

### 4.3 示例输入与输出

- [Data_Input](/C:/Users/super/capacity_optimizer/Data_Input)
- [output](/C:/Users/super/capacity_optimizer/output)

### 4.4 license 相关

- [licenses](/C:/Users/super/capacity_optimizer/licenses)
- [license_admin](/C:/Users/super/capacity_optimizer/license_admin)

## 5. 哪些目录是源码，哪些是运行产物

### 5.1 源码/文档目录

这些通常应该进入版本管理并作为交接主体：

- `app/`
- `build_support/`
- `packaging/`
- `tests/`
- `docs/`
- `requirements.txt`
- `README.md`
- `CapacityOptimizerLauncher.pyw`

### 5.2 运行产物/本地状态目录

这些通常是运行期或本地状态文件，接手时要知道它们存在，但不应把它们当成源码主体：

- `logs/`
- `output/`
- `dist/`
- `build/`
- `delivery_packages/`
- `launcher_settings.json`
- `workspace_manifest.json`

### 5.3 需谨慎对待的目录

以下目录可能包含本地历史、测试样例或中间状态，不要不加判断直接删掉：

- `licenses/`
- `license_admin/`
- `Data_Input/`
- `Tooling Control Panel/`

## 6. 本地运行方式

建议在项目根目录执行：

```powershell
cd C:\Users\super\capacity_optimizer
python -m pip install -r requirements.txt
python CapacityOptimizerLauncher.pyw
```

也可以直接双击：

- [CapacityOptimizerLauncher.pyw](/C:/Users/super/capacity_optimizer/CapacityOptimizerLauncher.pyw)

推荐操作顺序：

1. 打开 Launcher
2. 设置工作区路径
3. 点击 `Save Settings`
4. 点击 `Initialize Workspace`
5. 设置运行参数
6. 点击 `Run Optimization`

## 7. 打包方式

当前主交付方式是：

- `PyInstaller one-folder`

重点参考：

- [pyinstaller_onefolder_build.md](/C:/Users/super/capacity_optimizer/docs/pyinstaller_onefolder_build.md)

典型流程：

```powershell
cd C:\Users\super\capacity_optimizer
powershell -ExecutionPolicy Bypass -File packaging\build_onefolder.ps1
```

打包产物通常在：

- [dist](/C:/Users/super/capacity_optimizer/dist)
- [delivery_packages](/C:/Users/super/capacity_optimizer/delivery_packages)

## 8. license 方案

当前 license 是离线签名模式，基本流程是：

1. 客户端生成 machine fingerprint
2. 内部签发 `license.json`
3. 客户把 `license.json` 放入：
   - `licenses/active/license.json`
4. 启动时校验：
   - 签名
   - 到期时间
   - 机器绑定
   - 产品范围

相关说明文档：

- [CUSTOMER_LICENSE_QUICKSTART_CN.md](/C:/Users/super/capacity_optimizer/docs/CUSTOMER_LICENSE_QUICKSTART_CN.md)
- [INTERNAL_LICENSE_SOP_CN.md](/C:/Users/super/capacity_optimizer/docs/INTERNAL_LICENSE_SOP_CN.md)
- [IT_DEPLOYMENT_CHECKLIST_CN.md](/C:/Users/super/capacity_optimizer/docs/IT_DEPLOYMENT_CHECKLIST_CN.md)

## 9. 当前报表结构

当前主报告分三类：

### 9.1 单模式报告

- `ModeA`
- `ModeB`

planner 主要看的 sheet 已经过一轮精简和重排，重点页包括：

- `Dashboard`
- `Monthly_Trend`
- `Bottleneck`
- `WC_Heatmap`
- `Product_Risk`
- `Allocation_Detail`
- `Planner_Result_Summary`

### 9.2 模式对比报告

- `Summary of Mode A and Mode B`

当前已重点优化过的 Summary 页包括：

- `综合对比总览`
- `月度趋势对比`
- `瓶颈对比`

### 9.3 产能口径对比页

Summary 中还包含：

- `模式A产能口径总览`
- `模式A产能口径热力图`
- `模式B产能口径总览`
- `模式B产能口径热力图`

## 10. 当前项目接手时最值得注意的点

### 10.1 文档编码状态不完全统一

仓库中已有部分旧文档存在历史编码问题。继续维护时建议：

- 优先相信新生成或新重写的 UTF-8 文档
- 对旧乱码段落单独清理，不要在大改功能时顺手混改一大片文档

### 10.2 报表逻辑高度集中在 `app/output_writer.py`

这是当前项目最复杂、最容易引发连锁影响的文件。修改报表时建议：

- 先确认本次改动影响哪类 workbook
- 同步检查：
  - 单模式报告
  - Summary 报告
  - 产能口径对比报告

### 10.3 `_Dashboard_Helper` 这类辅助数据有“内部键”和“显示文字”两层

后续如果继续推进中英文化或公式筛选逻辑，要特别小心：

- 内部 helper / key 值尽量保持稳定 canonical 值
- 可见页面再做语言映射

否则容易出现公式匹配不到、整页显示 0 的问题。

## 11. 下一位开发者如何让 Codex 接手

最推荐的开场方式：

```text
先不要写代码。
先阅读以下文档并理解当前项目状态：
- README.md
- docs/PROJECT_HANDOFF_CN.md
- docs/TECHNICAL_REFERENCE_CN.md
- docs/developer_guide.md
- docs/runtime_directory_strategy.md
- docs/NEXT_PROJECT_REUSE_GUIDE_CN.md
- docs/NEW_PROJECT_KICKOFF_PROMPT_CN.md

然后输出：
1. 你对当前项目架构的理解
2. 当前已完成能力
3. 当前剩余开发事项
4. 你建议的下一步实施顺序
5. 开发前你还需要确认的风险点
```

如果是全新线程，强烈建议先走“先读文档、先规划、再改代码”的方式。

## 12. 推荐交给下一位开发者的文件清单

最少应交付：

- GitHub 仓库地址
- 当前建议接手分支/版本号
- 本文件：[PROJECT_HANDOFF_CN.md](/C:/Users/super/capacity_optimizer/docs/PROJECT_HANDOFF_CN.md)
- [TECHNICAL_REFERENCE_CN.md](/C:/Users/super/capacity_optimizer/docs/TECHNICAL_REFERENCE_CN.md)
- [NEXT_PROJECT_REUSE_GUIDE_CN.md](/C:/Users/super/capacity_optimizer/docs/NEXT_PROJECT_REUSE_GUIDE_CN.md)
- [NEW_PROJECT_KICKOFF_PROMPT_CN.md](/C:/Users/super/capacity_optimizer/docs/NEW_PROJECT_KICKOFF_PROMPT_CN.md)
- 一份代表性的 `Data_Input`
- 一份代表性的 `ModeA / ModeB / Summary` 输出样例

## 13. 建议的交接顺序

1. 先交 GitHub 仓库地址和当前版本说明
2. 再交本文件和相关技术文档
3. 再说明如何本地运行和打包
4. 最后再说明当前尚未完成或下一阶段优先项

这样接手的人不会只拿到代码，而是拿到：

- 项目目标
- 项目结构
- 运行方式
- 打包方式
- license 方式
- Codex 接手方式

这才是完整的“交接包”。
