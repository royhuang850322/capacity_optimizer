# Chemical Capacity Optimizer - IT 部署清单

当前版本使用 PySide6 桌面启动器作为用户入口。旧 Excel 控制工作簿 UI 已归档，不再作为部署验收项。

## 1. 部署目标

部署完成后，用户应能够：

1. 打开 `CapacityOptimizer.exe` 或源码模式下的 `CapacityOptimizerLauncher.pyw`。
2. 初始化工作区。
3. 放置或生成授权文件。
4. 在启动器中设置运行参数。
5. 点击 `Run Optimizer`。
6. 在输出目录查看 Excel 报告，在日志目录查看运行日志。

## 2. 环境要求

- Windows 10 / 11
- 本地磁盘可写
- Excel desktop application，用于查看输出报告
- 源码模式需要 Python；打包 exe 模式不要求用户手工运行 Python

## 3. 交付内容检查

源码模式至少包含：

```text
CapacityOptimizerLauncher.pyw
app\
Data_Input\
docs\
runtime\
requirements.txt
licenses\
```

打包模式至少包含：

```text
dist\CapacityOptimizer\CapacityOptimizer.exe
dist\CapacityOptimizer\_internal\resources\
```

不应再要求存在：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

该文件如需保留，仅应位于：

```text
Archive\legacy_excel_control_panel\
```

## 4. 安装步骤

### 源码模式

```powershell
cd <tool-root>
runtime\setup_requirements.bat
python CapacityOptimizerLauncher.pyw
```

### 打包模式

```powershell
dist\CapacityOptimizer\CapacityOptimizer.exe
```

## 5. 授权

推荐授权路径：

```text
licenses\active\license.json
```

机器指纹请求目录：

```text
licenses\requests\
```

## 6. 验收标准

- 可以打开启动器。
- 可以点击 `Initialize Workspace`。
- 可以生成机器指纹。
- 可以保存启动器设置。
- 可以点击 `Run Optimizer`。
- `output\` 中生成报告 workbook。
- `logs\` 中生成运行日志。
- 不需要打开 Excel 控制面板即可完成运行。
