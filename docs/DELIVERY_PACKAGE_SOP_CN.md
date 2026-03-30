# Chemical Capacity Optimizer - 客户交付包导出 SOP

本文档给 RSCP 内部使用。目标是把开发仓库稳定导出成一份干净的客户交付包。

---

## 1. 目标

导出的客户交付包应当：

- 保留客户运行工具所需的文件
- 排除内部开发和发证管理文件
- 自动生成一份新的控制工作簿
- 可选附带某个客户的 `license.json`

---

## 2. 导出脚本位置

- [export_customer_package.py](/C:/Users/super/capacity_optimizer/license_admin/export_customer_package.py)
- [open_delivery_exporter.bat](/C:/Users/super/capacity_optimizer/license_admin/open_delivery_exporter.bat)
- [delivery_exporter_ui.py](/C:/Users/super/capacity_optimizer/license_admin/delivery_exporter_ui.py)

如果不想手敲 PowerShell 命令，优先使用：

- 双击 `license_admin\open_delivery_exporter.bat`

它会打开一个本地 GUI 表单，你只需要填写：

- `Customer Name`
- `Destination Root`
- `Package Name`（可选）
- `License File`（可选）
- `Include demo data`
- `Overwrite`

然后点击 `Export Package` 即可。

---

## 3. 默认导出位置

默认输出到：

```text
delivery_packages\
```

例如：

```text
delivery_packages\capacity_optimizer_DuPont
```

该目录已被 `.gitignore` 忽略，不会进入 Git 仓库。

---

## 4. 导出命令

### 4.1 不带 license 的交付包

```powershell
python license_admin\export_customer_package.py `
  --customer-name "DuPont" `
  --overwrite
```

### 4.2 带指定 license 的交付包

```powershell
python license_admin\export_customer_package.py `
  --customer-name "DuPont" `
  --license-file "D:\RSCP_License_Admin\DuPont\capacity_optimizer\active\license.json" `
  --overwrite
```

### 4.3 不附带演示数据

```powershell
python license_admin\export_customer_package.py `
  --customer-name "DuPont" `
  --no-demo-data `
  --overwrite
```

---

## 5. 导出包内容

导出包会包含：

- `app\`
- `runtime\`
- `Tooling Control Panel\`
- `Data_Input\` 或空目录
- `output\`
- `licenses\active\`
- `licenses\requests\`
- `docs\CUSTOMER_LICENSE_QUICKSTART_CN.md`
- `docs\IT_DEPLOYMENT_CHECKLIST_CN.md`
- `requirements.txt`
- `LICENSE`
- 交付包专用 `README.md`
- `delivery_manifest.json`

不会包含：

- `tests\`
- `license_admin\`
- 内部发证 GUI
- 私钥
- 内部 SOP 文档
- Git 元数据

---

## 6. 控制工作簿

导出脚本不会直接复制你当前本机正在使用的 workbook。  
它会在交付包内自动重新生成：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

这样能避免把你本机保存过的参数、历史设置或临时路径带给客户。

---

## 7. 授权文件策略

推荐客户包里的授权位置：

```text
licenses\active\license.json
```

如果导出时未指定 `--license-file`：

- 交付包仍会创建 `licenses\active\`
- 但不会放入授权文件

如果导出时指定了 `--license-file`：

- 该文件会复制到交付包的：
  - `licenses\active\license.json`

---

## 8. 导出后检查项

每次导出完成后，建议至少检查：

1. 包内存在 `runtime\run_optimizer.bat`
2. 包内存在 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx`
3. 包内存在 `docs\CUSTOMER_LICENSE_QUICKSTART_CN.md`
4. 包内不存在 `license_admin\`
5. 包内不存在 `tests\`
6. 如果本次应附带授权，确认：
   - `licenses\active\license.json` 存在

---

## 9. 推荐发包流程

1. 在开发仓库完成代码修改
2. 跑测试
3. 如需正式授权，先完成发证
4. 运行交付导出脚本
5. 检查导出包结构
6. 压缩交付包并发送客户

---

## 10. 一句话版本

先在开发仓库里完成测试，再运行：

```powershell
python license_admin\export_customer_package.py --customer-name "DuPont" --overwrite
```

如果要把授权一起带给客户，再加：

```powershell
--license-file "...\license.json"
```
