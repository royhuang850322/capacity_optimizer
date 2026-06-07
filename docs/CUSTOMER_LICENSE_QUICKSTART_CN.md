# Chemical Capacity Optimizer - 客户授权使用说明

本文面向业务用户，说明如何在当前桌面启动器流程下放置授权、生成机器指纹并运行工具。

## 1. 当前入口

当前 UI 是桌面启动器：

```text
CapacityOptimizer.exe
```

源码模式可以打开：

```text
CapacityOptimizerLauncher.pyw
```

旧的 `Tooling Control Panel\Capacity_Optimizer_Control.xlsx` 已归档，不再作为客户运行入口。

## 2. 第一次使用

1. 将完整工具目录复制到本地磁盘。
2. 如果是源码模式，双击 `runtime\setup_requirements.bat` 安装依赖。
3. 打开 `CapacityOptimizerLauncher.pyw` 或打包后的 `CapacityOptimizer.exe`。
4. 点击 `Initialize Workspace`。
5. 将授权文件放到 `licenses\active\license.json`。
6. 在启动器中设置运行参数并点击 `Save Settings`。
7. 点击 `Run Optimizer`。

## 3. 试用授权

如果 RSCP 提供的是 `trial` 或 `unbound` 授权，通常不需要先生成机器指纹。直接把 `license.json` 放到：

```text
licenses\active\license.json
```

然后通过启动器运行。

## 4. 正式机绑授权

如果需要机器绑定授权：

1. 在启动器中点击 `Generate Machine Fingerprint`，或运行 `runtime\get_machine_fingerprint.bat`。
2. 从 `licenses\requests\` 中取最新的 `machine_fingerprint_*.json` 发给 RSCP。
3. 收到签名后的 `license.json`。
4. 将它放到 `licenses\active\license.json`。
5. 重新打开启动器并运行。

## 5. 常见问题

- 找不到授权：确认文件名必须是 `license.json`，并且位于 `licenses\active\`。
- 授权过期：联系 RSCP 重新签发。
- 机器不匹配：重新生成机器指纹并申请新授权。
- 运行失败：从启动器打开 `Log Folder`，把最新日志发给支持人员。
