# Chemical Capacity Optimizer - 用户手册

当前版本的正式用户入口是桌面启动器，不再是 Excel 控制工作簿。

推荐阅读：

- `docs\user_guide.md`
- `docs\desktop_launcher_usage.md`
- `docs\Capacity_Optimizer_v2.2.2_User_Guide_CN.docx`

## 快速流程

1. 打开 `CapacityOptimizer.exe`，或源码模式打开 `CapacityOptimizerLauncher.pyw`。
2. 点击 `Initialize Workspace`。
3. 放置 `licenses\active\license.json`。
4. 维护 `Data_Input` 下的输入文件。
5. 在启动器中设置 Run Mode、Scenario、Start Year、Start Month、Horizon Months、Output File Name 等参数。
6. 点击 `Save Settings`。
7. 点击 `Run Optimizer`。
8. 在 `output` 中查看 Excel 报告，在 `logs` 中查看运行日志。

## 旧 Excel 控制簿说明

历史版本使用：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

该入口已经归档到：

```text
Archive\legacy_excel_control_panel\
```

它只用于历史追溯和 legacy workbook 兼容测试，不作为当前业务用户入口。
