# Chemical Capacity Optimizer - IT 部署清单

本文档适用于当前版本的工具。

当前版本不是浏览器系统，不使用 Streamlit 页面。  
用户通过：

1. Excel 控制工作簿设置参数
2. Python 程序运行优化
3. Excel 结果工作簿查看报告

---

## 1. 部署目标

部署完成后，用户可以在 Windows 电脑上完成以下操作：

1. 打开 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. 在 `Control_Panel` 中设置输入路径、输出路径、Scenario、起始年月、Horizon、Run Mode 等参数
3. 双击 `run_optimizer.bat` 或执行 Python 命令运行
4. 在 `output\` 中查看结果 Excel

如果 `Run_Mode = Both`，还会额外生成：

- `Summary of Mode A and Mode B_YYYYMMDD_HHMMSS.xlsx`

---

## 2. IT 需要准备的环境

- Windows 10 / 11 或 Windows Server
- Python 3.10 及以上
- 本地磁盘可写
- 可以访问项目目录和输入数据目录

建议：

- CPU：4 核或以上
- 内存：8 GB 或以上
- 磁盘：至少 2 GB 可用空间

---

## 3. 项目目录建议

建议部署到：

```text
C:\Apps\capacity_optimizer
```

---

## 4. 需要交付给 IT 的内容

- 完整项目目录
- 输入数据目录
- 用户最终使用的控制工作簿路径

至少确认这些文件存在：

- `main.py`
- `create_template.py`
- `requirements.txt`
- `run_optimizer.bat`
- `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`

---

## 5. 安装步骤

### 5.1 复制项目

将项目目录复制到目标机器，例如：

```text
C:\Apps\capacity_optimizer
```

### 5.2 检查 Python

在 PowerShell 中执行：

```powershell
python --version
pip --version
```

### 5.3 创建虚拟环境

```powershell
cd C:\Apps\capacity_optimizer
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 5.4 安装依赖

```powershell
setup_requirements.bat
```

### 5.5 生成控制工作簿

如果项目里还没有最新控制工作簿，执行：

```powershell
python create_template.py
```

生成文件：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

---

## 6. 首次运行验证

### 6.1 编辑控制工作簿

打开：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

在 `Control_Panel` sheet 填写或确认：

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

推荐默认值：

- `Project_Root_Folder = ..`
- `Input_Load_Folder = Data_Input`
- `Input_Master_Folder = Data_Input`
- `Output_Folder = output`

### 6.2 命令行运行

```powershell
python main.py --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

### 6.3 批处理运行

也可以直接双击：

```text
run_optimizer.bat
```

它会：

1. 查找 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. 运行优化
3. 成功后打开 `output\` 目录

---

## 7. 验收标准

以下内容同时满足，即视为部署完成：

1. 可以打开 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. 可以成功运行 `python main.py --input-template "...Control.xlsx"`
3. `output\` 目录中能生成结果工作簿
4. 结果工作簿包含以下 sheet：

- `Dashboard`
- `Monthly_Trend`
- `Bottleneck`
- `WC_Heatmap`
- `Product_Risk`
- `Allocation_Detail`
- `Allocation_Summary`
- `Outsource_Summary`
- `Unmet_Summary`
- `WC_Load_Pct`
- `Binary_Feasibility`
- `Validation_Issues`
- `Run_Info`

如果运行模式为 `Both`，还应额外生成：

- `Summary of Mode A and Mode B_YYYYMMDD_HHMMSS.xlsx`

---

## 8. 日常使用方式

业务用户日常只需要：

1. 更新输入目录内的 CSV / Excel 数据
2. 打开控制工作簿调整参数
3. 双击 `run_optimizer.bat` 或运行命令
4. 打开输出结果 Excel

---

## 9. 常见问题

### 9.1 控制工作簿不存在

执行：

```powershell
python create_template.py
```

### 9.2 输入目录改了

直接到 `Control_Panel` 中修改：

- `Project_Root_Folder`
- `Input_Load_Folder`
- `Input_Master_Folder`
- `Output_Folder`

### 9.3 运行时提示验证错误

先修正输入数据。  
如果是演示或强制试跑，可将 `Skip_Validation_Errors` 改为 `Yes`。

### 9.4 结果文件没有生成

检查：

- `Output_Folder` 是否存在
- 当前用户是否有写权限
- 控制工作簿中的 `Output_FileName` 是否合法

---

## 10. 给 IT 的一句话版本

安装 Python 和依赖，保留项目目录与控制工作簿，用户通过 Excel 填参数，再运行：

```powershell
python main.py --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

结果会输出到控制工作簿指定的 `Output_Folder` 中。
