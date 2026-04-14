# Python 安装说明（客户版）

## 1. 什么时候需要看这份说明

如果运行工具时看到下面这类提示，说明当前电脑上的 Python 还没有正确安装或没有正确配置：

- `Python was not found`
- `Failed to generate machine fingerprint`
- `Run runtime\check_python_setup.bat first`

这不是工具本身损坏，而是当前电脑还不能正常运行 Python。

## 2. 安装原则

本工具客户侧统一按下面方式安装 Python：

- 只通过 **Windows Microsoft Store（微软应用商店）** 安装
- 不使用其他来源的 Python 安装包

## 3. 去哪里安装

请按下面步骤操作：

1. 打开 Windows 开始菜单
2. 打开 **Microsoft Store**
3. 在搜索框中输入：`Python`
4. 选择官方 Python 版本进行安装

建议：

- 选择稳定版本
- 建议安装 `Python 3.12` 或更新版本

## 4. 安装完成后怎么检查

安装完成后：

1. 关闭当前所有命令窗口
2. 重新打开一个新的 `Command Prompt` 或 `PowerShell`
3. 在命令窗口里输入：

```bash
python --version
```

如果能看到类似下面的结果，就说明 Python 已经安装成功：

```text
Python 3.12.x
```

或者：

```text
Python 3.13.x
```

## 5. 如果安装后仍然提示 Python not found

请继续执行：

```bash
runtime\check_python_setup.bat
```

这个脚本会帮助检查：

- 当前电脑是否已经安装 Python
- 是否是 Microsoft Store 的 Python 启动别名问题
- 是否需要自动修复 PATH

如果脚本检测到 Python 已经装好但没有正确加入 PATH，它会尝试自动修复。

## 6. 正确的工具安装顺序

当 Python 安装完成并能正常显示版本号后，请按下面顺序继续：

1. 运行：

```bash
runtime\check_python_setup.bat
```

2. 运行：

```bash
runtime\setup_requirements.bat
```

3. 如果需要申请正式机绑授权，再运行：

```bash
runtime\get_machine_fingerprint.bat
```

4. 收到 `license.json` 后，将它放到：

```text
licenses\active\license.json
```

5. 最后运行：

```bash
runtime\run_optimizer.bat
```

## 7. 如果客户已经安装过 Python，但工具仍然报错

这通常是以下原因之一：

- 安装后没有重新打开命令窗口
- Windows 还在使用旧的 Python 启动别名
- PATH 没有正确配置

请优先运行：

```bash
runtime\check_python_setup.bat
```

不要先手动修改系统配置。

## 8. 给客户的最短说明

如果客户只需要最短处理步骤，可以直接发下面这段：

1. 打开 **Microsoft Store**
2. 搜索并安装 **Python**
3. 安装完成后重新打开命令窗口
4. 在工具目录运行：

```bash
runtime\check_python_setup.bat
runtime\setup_requirements.bat
```

如果仍有问题，请把报错截图发回。
