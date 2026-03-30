# Chemical Capacity Optimizer - 客户授权使用说明

本文档给客户使用。目标是让客户从拿到工具开始，可以一步一步完成：

1. 试用版运行
2. 正式版切换
3. 续期
4. 换电脑后重新授权

---

## 1. 你会收到什么

通常你会从 RSCP 收到：

- 完整工具文件夹
- 一个 `license.json`

有两种情况：

1. 试用版
   - RSCP 直接提供 `license.json`
   - 不需要先提供机器信息

2. 正式版
   - 需要先生成本机机器指纹
   - 再由 RSCP 提供 `license.json`

---

## 2. 第一次在新电脑上使用工具

### Step 1：复制整个工具文件夹

把整个 `capacity_optimizer` 文件夹复制到你的电脑，例如：

```text
D:\capacity_optimizer
```

不要只复制 Excel 文件。

### Step 2：安装依赖

双击项目根目录里的：

```text
setup_requirements.bat
```

等待安装完成。

### Step 3：放置授权文件

把 RSCP 提供的：

```text
license.json
```

放到项目根目录，也就是和这些文件同级的位置：

- `main.py`
- `run_optimizer.bat`
- `setup_requirements.bat`

例如：

```text
D:\capacity_optimizer\license.json
```

### Step 4：打开控制工作簿

打开：

```text
Tooling Control Panel\Capacity_Optimizer_Control.xlsx
```

在 `Control_Panel` 里检查路径设置。推荐默认值：

- `Project_Root_Folder = ..`
- `Input_Load_Folder = Data_Input`
- `Input_Master_Folder = Data_Input`
- `Output_Folder = output`

### Step 5：保存后运行

先按 `Ctrl+S` 保存，再双击：

```text
run_optimizer.bat
```

成功后，工具会自动打开 `output` 文件夹。

---

## 3. 试用版怎么运行

如果 RSCP 提供的是试用版：

- 授权类型通常是 `trial`
- 绑定方式通常是 `unbound`

这表示：

- 你不需要先生成机器指纹
- 只要 `license.json` 在项目根目录，就可以直接运行

试用版的操作步骤就是：

1. 复制工具文件夹
2. 运行 `setup_requirements.bat`
3. 放入 RSCP 提供的 `license.json`
4. 打开控制工作簿
5. 保存后运行 `run_optimizer.bat`

---

## 4. 正式版怎么运行

如果 RSCP 要求提供机器信息，说明你使用的是正式版授权。

### Step 1：生成机器指纹

双击：

```text
get_machine_fingerprint.bat
```

运行后会在项目根目录生成：

```text
machine_fingerprint.json
```

### Step 2：把机器指纹发给 RSCP

把 `machine_fingerprint.json` 发给 RSCP。

### Step 3：接收正式版授权文件

RSCP 会返回正式版：

```text
license.json
```

### Step 4：放回项目根目录

把新的 `license.json` 放到项目根目录，然后运行：

```text
run_optimizer.bat
```

---

## 5. 试用版升级成正式版

如果你已经在使用试用版，后面转正式版时，不需要重新安装工具。

只需要：

1. 运行 `get_machine_fingerprint.bat`
2. 把 `machine_fingerprint.json` 发给 RSCP
3. 收到新的正式版 `license.json`
4. 用新的 `license.json` 替换旧的试用版 `license.json`
5. 重新运行工具

---

## 6. 后续续期怎么做

如果只是授权到期，但电脑没有变：

1. 向 RSCP 申请新的 `license.json`
2. 用新的 `license.json` 替换项目根目录里的旧文件
3. 重新运行工具

一般情况下：

- 不需要重新安装工具
- 不需要重新安装依赖
- 不需要重新生成机器指纹

---

## 7. 换电脑怎么做

如果你更换了电脑，正式版授权需要重新绑定新机器。

在新电脑上：

1. 复制整个工具文件夹
2. 运行 `setup_requirements.bat`
3. 运行 `get_machine_fingerprint.bat`
4. 把新的 `machine_fingerprint.json` 发给 RSCP
5. 收到新的 `license.json`
6. 把新的 `license.json` 放到新电脑的项目根目录
7. 运行工具

注意：

- 旧电脑的正式授权文件通常不能直接用于新电脑

---

## 8. 什么时候工具会停止运行

如果出现下面任一情况，工具会停止：

- 项目根目录没有 `license.json`
- `license.json` 被修改或无效
- 授权已过期
- 正式版授权绑定的机器与当前电脑不匹配

---

## 9. 常见问题

### 9.1 没有 `license.json`

联系 RSCP 获取授权文件。  
如果是正式版，请先运行：

```text
get_machine_fingerprint.bat
```

### 9.2 不知道 `license.json` 放哪里

放在项目根目录，也就是和这些文件同级：

- `main.py`
- `run_optimizer.bat`
- `setup_requirements.bat`

### 9.3 点了 `Run Optimizer` 不能运行

先检查：

- 是否已经运行过 `setup_requirements.bat`
- 项目根目录里是否有 `license.json`
- 控制工作簿是否已经保存

### 9.4 输出文件没有生成

检查：

- `Output_Folder` 是否正确
- 输入路径是否正确
- 授权是否仍然有效

---

## 10. 联系 RSCP 时建议提供的信息

如果需要授权或排错，请提供：

- 你的公司名称
- 当前使用的是试用版还是正式版
- `machine_fingerprint.json`（正式版时）
- 错误截图
- 当前 `license.json` 的到期日期

