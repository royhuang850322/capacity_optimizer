# Chemical Capacity Optimizer 使用手册

版本：v1.1.2  
适用对象：客户端用户、内部实施人员、开发维护人员

---

## 1. 文档目的

本文档用于完整说明 `Chemical Capacity Optimizer` 的使用、部署、授权、交付和维护流程。  
文档按两大对象组织：

- 客户端 / 最终用户：如何拿到工具后完成安装、授权、运行、查看结果和后续续期
- 开发端 / 内部运维端：如何维护代码、签发 license、导出客户交付包、发布版本

---

## 2. 工具简介

`Chemical Capacity Optimizer` 是一个以 Excel 为控制界面的产能优化工具，当前采用以下工作方式：

- 输入数据：CSV / Excel
- 运算逻辑：Python + OR-Tools
- 操作入口：Excel 控制工作簿
- 输出结果：Excel 报告工作簿

当前主流程已经不再依赖网页或 Streamlit。

---

## 3. 当前目录结构说明

项目根目录的主要结构如下：

```text
capacity_optimizer/
|-- app/                         核心 Python 程序
|-- runtime/                     客户侧运行入口脚本
|-- Tooling Control Panel/       Excel 控制工作簿
|-- Data_Input/                  演示输入数据
|-- output/                      输出结果目录
|-- licenses/                    客户运行时授权目录
|   |-- active/
|   `-- requests/
|-- docs/                        文档
|-- license_admin/               内部 license 管理工具
|-- tests/                       测试代码
|-- README.md
|-- requirements.txt
`-- LICENSE
```

---

## 4. 客户端使用说明

### 4.1 客户收到工具后需要看到的文件

客户侧通常只需要关心以下目录或文件：

- `runtime\setup_requirements.bat`
- `runtime\get_machine_fingerprint.bat`
- `runtime\run_optimizer.bat`
- `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
- `Data_Input\`
- `output\`
- `licenses\active\license.json`

客户不需要使用：

- `license_admin\`
- `tests\`
- 内部私钥
- 内部发证工具

---

### 4.2 客户端首次部署步骤

#### Step 1：复制整个工具文件夹

把整个 `capacity_optimizer` 文件夹复制到客户电脑，例如：

```text
D:\capacity_optimizer
```

不要只复制 Excel 文件。

#### Step 2：安装 Python

确认目标电脑已经安装 Python，并且以下命令可执行：

```powershell
python --version
```

#### Step 3：安装依赖

双击：

```text
runtime\setup_requirements.bat
```

这一步会自动安装运行工具所需的 Python 依赖。

#### Step 4：准备授权文件

把签发好的 `license.json` 放到：

```text
licenses\active\license.json
```

当前程序也兼容旧版根目录 `license.json`，但推荐统一使用 `licenses\active\license.json`。

#### Step 5：打开控制工作簿

打开：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

第一次打开时，先查看：

- `Deployment_Steps`
- `Instructions`

然后在 `Control_Panel` 中填写或确认参数。

#### Step 6：保存 Excel 后运行

修改参数后，请先保存 Excel，再运行：

```text
runtime\run_optimizer.bat
```

也可以使用命令行：

```powershell
python -m app.main --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

---

### 4.3 客户端授权说明

当前支持两类授权：

#### 1. Trial / Unbound

特点：

- 试用版
- 不绑定机器
- 适合短期演示或 POC

客户侧流程：

1. 从 RSCP 收到 `license.json`
2. 放入 `licenses\active\license.json`
3. 直接运行工具

#### 2. Machine-Locked

特点：

- 正式版
- 绑定某一台机器
- 更适合正式交付

客户侧流程：

1. 双击 `runtime\get_machine_fingerprint.bat`
2. 程序会在 `licenses\requests\` 下生成机器指纹文件
3. 把该文件发给 RSCP
4. 收到 RSCP 签发的 `license.json`
5. 放入 `licenses\active\license.json`
6. 再运行工具

---

### 4.4 机器指纹文件说明

运行：

```text
runtime\get_machine_fingerprint.bat
```

后，会在目录下生成类似文件：

```text
licenses\requests\machine_fingerprint_<MachineLabel>_YYYYMMDD_HHMMSS.json
```

这个文件用于向 RSCP 申请正式版机绑授权。

---

### 4.5 Control_Panel 关键参数说明

#### 一、迁移设置区

##### `Project_Root_Folder`

表示项目根目录路径。推荐保持：

```text
..
```

因为控制工作簿位于 `Tooling Control Panel\` 内，`..` 表示上一级，也就是整个工具根目录。

##### `Input_Load_Folder`

表示 planner 负荷文件目录。推荐：

```text
Data_Input
```

##### `Input_Master_Folder`

表示主数据目录。推荐：

```text
Data_Input
```

##### `Output_Folder`

表示输出结果目录。推荐：

```text
output
```

---

#### 二、运行设置区

##### `Scenario_Name`

场景名称，用于筛选或标识当前运行场景。

##### `Start_Year`

运行起始年份。

##### `Start_Month_Num`

运行起始月份，填 1 到 12。

##### `Horizon_Months`

向后计算多少个月。例如：

- `12` 表示 12 个月
- `24` 表示 24 个月

##### `Run_Mode`

可选：

- `ModeA`
- `ModeB`
- `Both`

说明：

- `ModeA`：只考虑内部产能
- `ModeB`：考虑 routing、内部分配、外包和 unmet
- `Both`：同时跑 `ModeA` 和 `ModeB`，并生成比较报告

##### `Direct_Mode`

建议保持：

```text
Yes
```

表示直接从文件夹读取 planner / master 数据。

##### `Verbose`

是否在命令窗口打印更多过程信息。

建议：

- 日常运行：`No`
- 排错时：`Yes`

##### `Skip_Validation_Errors`

是否在校验发现错误时仍然继续运行。

建议：

- 正常使用：`No`
- 仅在明确知道风险时，才临时改为 `Yes`

##### `Run_Timestamp`

运行时由程序自动写入，不需要手工填写。

##### `Notes`

自由备注，可空。

---

### 4.6 输入文件要求

#### Planner 文件

文件名支持：

- `planner1_load`
- `planner2_load`
- `planner3_load`
- `planner4_load`
- `planner5_load`
- `planner6_load`

支持扩展名：

- `.csv`
- `.xlsx`
- `.xls`

主要字段：

- `Month`
- `PlannerName`
- `Product`
- `ProductFamily`
- `Plant`
- `Forecast_Tons`

#### Capacity Master

必须提供：

- `master_capacity.csv`
  或对应 Excel 版本

#### Routing Master

`ModeB` 下必须提供：

- `alternative_routing`
  或
- `master_routing`

---

### 4.7 输出结果说明

每次运行都会在 `Output_Folder` 下生成时间戳文件。

#### 单模式运行

会生成：

- `capacity_result_ModeA_...xlsx`
或
- `capacity_result_ModeB_...xlsx`

#### `Run_Mode = Both`

会额外生成：

- `Summary of Mode A and Mode B_YYYYMMDD_HHMMSS.xlsx`

#### 常见 sheet

- `Dashboard`
- `Monthly_Trend`
- `Bottleneck`
- `WC_Heatmap`
- `Product_Risk`
- `Allocation_Detail`
- `Allocation_Summary`
- `Outsource_Summary`
- `Unmet_Summary`
- `Planner_Result_Summary`
- `Planner_Product_Month`
- `Run_Info`

---

### 4.8 客户端常见问题

#### 问题 1：运行时报 license not found

检查：

- `licenses\active\license.json` 是否存在
- 文件名是否正确
- 是否误放到别的目录

#### 问题 2：运行时报 license expired

说明授权已过期，需要联系 RSCP 续期并替换新的 `license.json`。

#### 问题 3：运行时报 machine does not match

说明当前 `license.json` 不是给这台电脑签发的。  
需要重新运行 `runtime\get_machine_fingerprint.bat`，并联系 RSCP 重签授权。

#### 问题 4：运行时报 openpyxl / pandas / package missing

重新执行：

```text
runtime\setup_requirements.bat
```

#### 问题 5：修改了 Excel 里的路径但程序仍读旧值

需要先在 Excel 中按 `Ctrl+S` 保存，再运行 `runtime\run_optimizer.bat`。

---

### 4.9 客户端后续授权维护

#### 续期

客户只需要：

1. 收到新的 `license.json`
2. 覆盖 `licenses\active\license.json`
3. 重新运行工具

#### 换电脑

客户需要：

1. 在新电脑复制整个工具文件夹
2. 运行 `runtime\setup_requirements.bat`
3. 运行 `runtime\get_machine_fingerprint.bat`
4. 把新生成的机器指纹文件发给 RSCP
5. 用新的正式版 `license.json` 替换旧授权

---

## 5. 开发端 / 内部维护说明

### 5.1 开发仓库和客户交付包的区别

当前仓库是开发仓库，包含：

- 核心程序
- 测试
- 文档
- 内部 license 管理工具
- 内部交付打包工具

客户交付包则只应包含客户真正需要的运行内容，不应包含：

- `tests\`
- `license_admin\private_keys\`
- 内部 SOP
- 内部发证工具

---

### 5.2 开发环境常用命令

#### 运行测试

```powershell
python -m unittest discover -s tests -v
```

#### 编译检查

```powershell
python -m py_compile app\main.py
```

#### 生成控制工作簿

```powershell
python -m app.create_template
```

#### 手工运行优化器

```powershell
python -m app.main --input-template "Tooling Control Panel\Capacity_Optimizer_Control.xlsx"
```

---

### 5.3 内部 license 管理目录标准

内部推荐统一使用：

```text
D:\RSCP_License_Admin\<CustomerName>\capacity_optimizer\
```

目录结构建议：

```text
requests\
issued\
active\
archive\
notes\
```

说明：

- `requests`：客户回传的机器指纹申请文件
- `issued`：签发出去的授权副本
- `active`：当前生效授权副本
- `archive`：历史授权
- `notes`：备注或台账

---

### 5.4 内部发证方式

#### 方式 1：GUI 方式

双击：

```text
license_admin\open_license_generator.bat
```

然后在 GUI 中填写：

- Customer Name
- Customer ID
- License ID
- License Type
- Issue Date
- Expiry Date
- Binding Mode
- Machine Fingerprint
- Machine Label
- Note

生成后会自动把文件写入内部 license 仓库，并更新对应客户的 `active\license.json`。

#### 方式 2：CLI 方式

试用版：

```powershell
python license_admin\license_tools\generate_trial_license.py ...
```

正式版：

```powershell
python license_admin\license_tools\generate_license.py ...
```

---

### 5.5 客户交付包导出方式

#### 方式 1：GUI 方式

双击：

```text
license_admin\open_delivery_exporter.bat
```

填写：

- `Customer Name`
- `Destination Root`
- `Package Name`（可选）
- `License File`（可选）
- 是否带演示数据
- 是否覆盖已有包

点击 `Export Package` 即可。

#### 方式 2：CLI 方式

```powershell
python license_admin\export_customer_package.py --customer-name "DuPont" --overwrite
```

如果要附带授权文件：

```powershell
python license_admin\export_customer_package.py `
  --customer-name "DuPont" `
  --license-file "D:\RSCP_License_Admin\DuPont\capacity_optimizer\active\license.json" `
  --overwrite
```

---

### 5.6 交付包内容说明

导出的客户交付包会包含：

- `app\`
- `runtime\`
- `Tooling Control Panel\`
- `docs\CUSTOMER_LICENSE_QUICKSTART_CN.md`
- `docs\IT_DEPLOYMENT_CHECKLIST_CN.md`
- `licenses\active\`
- `licenses\requests\`
- `output\`
- `README.md`
- `requirements.txt`
- `LICENSE`

不会包含：

- `tests\`
- `license_admin\`
- 私钥
- 内部 SOP

---

### 5.7 Git 与 GitHub 版本管理

#### 常规提交流程

```powershell
git status
git add .
git commit -m "feat: your change summary"
git push
```

#### 正式版本发布

```powershell
git tag v1.1.2
git push origin v1.1.2
```

#### 建议的提交前检查

```powershell
python -m unittest discover -s tests -v
```

---

### 5.8 推荐的版本号规则

- `v1.1.x`
  小功能新增、文档完善、交付流程增强
- `v1.2.x`
  中等级功能增强，但不破坏主流程
- `v2.0.0`
  输入结构、输出结构或主流程发生重大变化

---

### 5.9 开发端常见注意事项

#### 1. 私钥不要进 Git

私钥只应保留在内部受控目录，不应提交到 GitHub。

#### 2. 客户包不要整包发开发仓库

应当使用导出器生成客户交付包，而不是手工挑文件。

#### 3. 控制工作簿路径改了以后要重生成

如果脚本入口、目录结构或说明文字变化较大，建议重新运行：

```powershell
python -m app.create_template
```

#### 4. 修改批处理或路径后，要做真实 smoke test

至少验证：

- `runtime\setup_requirements.bat`
- `runtime\run_optimizer.bat`
- `runtime\get_machine_fingerprint.bat`
- `license_admin\open_license_generator.bat`
- `license_admin\open_delivery_exporter.bat`

---

## 6. 推荐日常工作流程

### 6.1 发试用版

1. 用内部 GUI 生成 `trial / unbound` license
2. 用交付导出器生成客户包
3. 可选择把 `license.json` 一起带入交付包
4. 发给客户

### 6.2 发正式版

1. 客户先提供机器指纹
2. 内部生成 `machine_locked` license
3. 用交付导出器导出正式交付包
4. 把授权一起带入客户包
5. 发给客户

### 6.3 续期

1. 在内部生成新的 license
2. 直接把新的 `license.json` 发给客户
3. 客户覆盖 `licenses\active\license.json`

### 6.4 客户换电脑

1. 让客户在新电脑运行机器指纹脚本
2. 收到新的 fingerprint 文件
3. 重新签发正式版 license
4. 让客户替换为新的 `license.json`

---

## 7. 文档索引

如需更细的操作说明，可参考：

- `README.md`
- `docs\CUSTOMER_LICENSE_QUICKSTART_CN.md`
- `docs\IT_DEPLOYMENT_CHECKLIST_CN.md`
- `docs\INTERNAL_LICENSE_SOP_CN.md`
- `docs\DELIVERY_PACKAGE_SOP_CN.md`

---

## 8. 结语

当前版本的 `Chemical Capacity Optimizer` 已经形成完整的工作闭环：

- 客户可通过 Excel 控制工作簿运行工具
- 内部可通过 GUI 方式签发 license
- 内部可通过 GUI 或 CLI 导出干净的客户交付包
- 授权、交付、运行、续期、换机都已有明确流程

建议今后继续保持：

- 客户运行文件和内部管理文件分离
- 所有正式交付包由导出器生成
- 所有正式版本通过 Git tag 管理
