# Chemical Capacity Optimizer - IT 部署清单

本文档适用于当前版本的工具。

当前版本不是浏览器系统，不使用 Streamlit 页面。用户通过：

1. Excel 控制工作簿设置参数
2. Python 程序运行优化
3. Excel 结果工作簿查看报告
4. `license.json` 授权文件控制是否可运行

---

## 1. 部署目标

部署完成后，用户可以在 Windows 电脑上完成以下操作：

1. 打开 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. 在 `Control_Panel` 中设置输入路径、输出路径、Scenario、起始年月、Horizon、Run Mode 等参数
3. 双击 `runtime\run_optimizer.bat` 或执行 Python 命令运行
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
- 客户对应的 `license.json`

至少确认这些内容存在：

- `app\main.py`
- `app\create_template.py`
- `requirements.txt`
- `runtime\setup_requirements.bat`
- `runtime\get_machine_fingerprint.bat`
- `runtime\run_optimizer.bat`
- `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`

推荐授权位置：

- `licenses\active\license.json`

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
runtime\setup_requirements.bat
```

### 5.5 生成控制工作簿

如果项目里还没有最新控制工作簿，执行：

```powershell
python -m app.create_template
```

生成文件：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

### 5.6 获取授权文件

当前支持两种授权方式：

1. 试用版 / `unbound`
   - RSCP 直接签发 `license.json`
   - 不需要机器指纹

2. 正式版 / `machine_locked`
   - 先生成机器指纹
   - 再由 RSCP 针对指定电脑签发 `license.json`

### 5.7 生成机器指纹

如果要给这台电脑发正式授权，先执行：

```powershell
runtime\get_machine_fingerprint.bat
```

它会在下面目录生成带时间戳的请求文件：

```text
licenses\requests\
```

把这个文件发给 RSCP，换取当前电脑专用的 `license.json`。

### 5.8 放置授权文件

把 RSCP 返回的 `license.json` 放到：

```text
licenses\active\license.json
```

兼容说明：

- 当前程序仍兼容旧的项目根目录 `license.json`
- 但 IT 部署时建议统一使用 `licenses\active\license.json`

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
python -m app.main --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

### 6.3 批处理运行

也可以直接双击：

```text
runtime\run_optimizer.bat
```

它会：

1. 查找 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. 检查依赖和授权文件
3. 运行优化
4. 成功后打开 `output\` 目录

如果缺少正式授权文件，先运行：

```powershell
runtime\get_machine_fingerprint.bat
```

把 `licenses\requests\` 下生成的机器指纹文件发给 RSCP，再把返回的 `license.json` 放到 `licenses\active\license.json`。

---

## 7. 验收标准

以下内容同时满足，即视为部署完成：

1. 可以打开 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
2. `licenses\active\license.json` 或兼容根目录 `license.json` 存在且有效
3. 可以成功运行 `python -m app.main --input-template "...Control.xlsx"`
4. `output\` 目录中能生成结果工作簿
5. 结果工作簿 `Run_Info` 中可看到授权信息：
   - `License_Status`
   - `License_ID`
   - `Licensed_To`
   - `License_Expiry`
6. 结果工作簿包含以下 sheet：

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

1. 保持有效的 `license.json`
2. 更新输入目录内的 CSV / Excel 数据
3. 打开控制工作簿调整参数
4. 双击 `runtime\run_optimizer.bat` 或运行命令
5. 打开输出结果 Excel

---

## 9. 常见问题

### 9.1 控制工作簿不存在

执行：

```powershell
python -m app.create_template
```

### 9.2 缺少授权文件

如果是试用版，直接向 RSCP 索取 `trial / unbound` 的 `license.json`。  
如果是正式版，执行：

```powershell
runtime\get_machine_fingerprint.bat
```

把 `licenses\requests\` 下生成的机器指纹文件发给 RSCP，并把返回的 `license.json` 放到 `licenses\active\license.json`。

### 9.3 输入目录改了

直接到 `Control_Panel` 中修改：

- `Project_Root_Folder`
- `Input_Load_Folder`
- `Input_Master_Folder`
- `Output_Folder`

### 9.4 运行时提示验证错误

先修正输入数据。  
如果是演示或强制试跑，可将 `Skip_Validation_Errors` 改为 `Yes`。

### 9.5 结果文件没有生成

检查：

- `Output_Folder` 是否存在
- 当前用户是否有写权限
- 控制工作簿中的 `Output_FileName` 是否合法
- `license.json` 是否存在且未过期

---

## 10. 给 IT 的一句话版本

安装 Python 和依赖，运行 `runtime\get_machine_fingerprint.bat` 申请授权，把 `license.json` 放到 `licenses\active\license.json`，然后运行：

```powershell
python -m app.main --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

结果会输出到控制工作簿指定的 `Output_Folder` 中。
