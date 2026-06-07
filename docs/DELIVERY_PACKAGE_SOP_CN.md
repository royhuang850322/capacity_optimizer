# Capacity Optimizer - 交付包 SOP

当前交付包采用 launcher-first 流程，不再生成或交付 Excel 控制工作簿 UI。

## 1. 交付包应包含

```text
CapacityOptimizerLauncher.pyw
app\
runtime\
Data_Input\
docs\
licenses\
output\
logs\
README.md
delivery_manifest.json
```

## 2. 交付包不应包含

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
license_admin\
tests\
```

## 3. 客户首次运行

1. 源码模式先运行 `runtime\setup_requirements.bat`。
2. 打开 `CapacityOptimizerLauncher.pyw`。
3. 点击 `Initialize Workspace`。
4. 放置 `licenses\active\license.json`。
5. 在启动器中设置参数。
6. 点击 `Run Optimizer`。

## 4. 内部导出命令

使用 `license_admin.export_customer_package` 生成源码模式交付包。该导出逻辑会复制 launcher、runtime、Data_Input、客户文档和授权目录，但不会生成旧 Excel 控制簿。

## 5. 验收

- 包内存在 `CapacityOptimizerLauncher.pyw`。
- 包内不存在 `Tooling Control Panel`。
- 包内存在 `licenses\active` 和 `licenses\requests`。
- 包内存在 `logs` 和 `output`。
- README 指向桌面启动器流程。
