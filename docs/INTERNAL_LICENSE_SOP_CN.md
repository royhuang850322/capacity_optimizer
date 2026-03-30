# Chemical Capacity Optimizer - 内部 License 发放 SOP

本文档给 RSCP 内部使用。目标是统一：

1. 试用版发放
2. 正式版发放
3. 试用转正式
4. 续期
5. 换电脑

当前 v1 授权方案基于：

- `license.json`
- `Ed25519` 数字签名
- `expiry_date`
- `binding_mode`
- 可选机器绑定 `machine_fingerprint`

内部日常推荐入口：

- [open_license_generator.bat](/C:/Users/super/capacity_optimizer/license_admin/open_license_generator.bat)
- [客户交付包导出 SOP](/C:/Users/super/capacity_optimizer/docs/DELIVERY_PACKAGE_SOP_CN.md)
- [open_delivery_exporter.bat](/C:/Users/super/capacity_optimizer/license_admin/open_delivery_exporter.bat)

双击后会打开本地 GUI 表单，可以直接填写参数并生成 `license.json`。命令行脚本仍然保留，用于备用或批量处理。

内部统一目录建议：

```text
D:\RSCP_License_Admin\<CustomerName>\capacity_optimizer\
```

并固定拆分为：

- `requests`
- `issued`
- `active`
- `archive`
- `notes`

---

## 1. 当前支持的授权类型

### 1.1 试用版

建议配置：

- `license_type = trial`
- `binding_mode = unbound`

特点：

- 不需要客户先回传机器指纹
- 交付速度快
- 适合 Demo / POC / 短期试用

### 1.2 正式版

建议配置：

- `license_type = commercial`
- `binding_mode = machine_locked`

特点：

- 绑定指定电脑
- 更适合正式交付

---

## 2. 内部要保管的内容

必须只保留在 RSCP 内部，不发给客户：

- 私钥文件
  - 例如：`license_admin\private_keys\license_signing_ed25519_private.pem`
- 发证脚本
  - [generate_license.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/generate_license.py)
  - [generate_trial_license.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/generate_trial_license.py)
  - [license_generator_ui.py](/C:/Users/super/capacity_optimizer/license_admin/license_tools/license_generator_ui.py)

可以发给客户的只有：

- `license.json`

---

## 3. 客户首次试用的标准流程

### 3.1 推荐场景

适用于：

- 演示
- 短期测试
- 先让客户跑起来

### 3.2 内部操作步骤

1. 确认客户名称和客户编号
2. 设定试用天数，例如 7 天、14 天、30 天
3. 运行试用版签发脚本
4. 把生成的 `license.json` 发给客户

如果你不想手敲命令，优先使用：

- [open_license_generator.bat](/C:/Users/super/capacity_optimizer/license_admin/open_license_generator.bat)

示例命令：

```powershell
python license_admin\license_tools\generate_trial_license.py `
  --private-key license_admin\private_keys\license_signing_ed25519_private.pem `
  --license-id LIC-TRIAL-2026-0001 `
  --customer-name "ABC Chemical" `
  --customer-id "ABC-001" `
  --days-valid 14 `
  --note "14-day trial"
```

### 3.3 对客户的说明

告诉客户：

1. 复制整个工具文件夹
2. 运行 `runtime\setup_requirements.bat`
3. 把 `license.json` 放到 `licenses\active\license.json`
4. 运行工具

---

## 4. 正式版首次交付流程

### 4.1 客户侧动作

客户需要先在目标电脑上运行：

```text
runtime\get_machine_fingerprint.bat
```

然后把生成的：

```text
licenses\requests\machine_fingerprint_*.json
```

发回 RSCP。

### 4.2 内部操作步骤

1. 读取客户回传的 `machine_fingerprint.json`
2. 确认：
   - `customer_name`
   - `customer_id`
   - `machine_fingerprint`
   - `machine_label`
   - 生效日期
   - 到期日期
3. 运行正式版签发脚本
4. 把生成的 `license.json` 发给客户

如果你使用 GUI：

1. 双击 `license_admin\open_license_generator.bat`
2. 选择 `Manual / Custom`
3. 选择 `binding_mode = machine_locked`
4. 填入客户信息和日期
5. 导入或粘贴 `machine_fingerprint`
6. 点击 `Generate License`

生成后建议检查：

- `issued\<LicenseID>.json`
- `active\license.json`

都已经更新到对应客户目录下

示例命令：

```powershell
python license_admin\license_tools\generate_license.py `
  --private-key license_admin\private_keys\license_signing_ed25519_private.pem `
  --license-id LIC-COMM-2026-0001 `
  --license-type commercial `
  --customer-name "ABC Chemical" `
  --customer-id "ABC-001" `
  --issue-date 2026-03-29 `
  --expiry-date 2027-03-28 `
  --binding-mode machine_locked `
  --machine-fingerprint "sha256:..." `
  --machine-label "ABC-LAPTOP-01" `
  --note "Commercial annual license"
```

### 4.3 对客户的说明

告诉客户：

1. 把 `license.json` 放到 `licenses\active\license.json`
2. 运行 `runtime\run_optimizer.bat`

---

## 5. 试用版转正式版流程

这是最推荐的实际交付路径。

### 5.1 客户已经在跑试用版

此时客户已经有：

- 工具文件夹
- 已安装好的依赖
- 当前试用版 `license.json`

### 5.2 客户需要做的事

1. 运行 `runtime\get_machine_fingerprint.bat`
2. 把 `licenses\requests\` 下生成的机器指纹文件发给 RSCP

### 5.3 RSCP 内部要做的事

1. 签发 `machine_locked` 正式版 `license.json`
2. 发给客户

### 5.4 客户最终动作

1. 用新的正式版 `license.json` 替换 `licenses\active\license.json`
2. 重新运行工具

注意：

- 不需要重发整包
- 不需要重新安装依赖

---

## 6. 续期流程

### 6.1 试用版续期

适用于客户还在评估，需要延长试用。

内部做法：

1. 重新生成一个新的 `trial / unbound` `license.json`
2. 更新：
   - `license_id`
   - `issue_date`
   - `expiry_date`
3. 发给客户覆盖旧文件

### 6.2 正式版续期

适用于：

- 客户已正式使用
- 电脑没有变化

内部做法：

1. 使用原有 `machine_fingerprint`
2. 重新生成新的 `machine_locked` `license.json`
3. 更新：
   - `license_id`
   - `issue_date`
   - `expiry_date`
4. 发给客户覆盖旧文件

一般不需要让客户重新回传机器指纹，除非：

- 内部记录不完整
- 你怀疑客户已换电脑

---

## 7. 换电脑流程

### 7.1 适用情况

- 客户电脑更换
- 客户系统重装后需要重新部署
- 工具迁移到另一台设备

### 7.2 客户要做的事

1. 在新电脑复制工具文件夹
2. 运行 `runtime\setup_requirements.bat`
3. 运行 `runtime\get_machine_fingerprint.bat`
4. 把新的 `licenses\requests\` 里的机器指纹文件发给 RSCP

### 7.3 RSCP 内部要做的事

1. 用新机器指纹重新签发正式版 `license.json`
2. 发给客户

### 7.4 注意事项

- 旧电脑的正式授权一般不能直接用于新电脑
- 这是正常设计，不属于工具异常

---

## 8. 建议的 license 编号规则

建议至少保证唯一且可读。

可以参考：

- `LIC-TRIAL-2026-0001`
- `LIC-COMM-2026-0001`
- `LIC-RENEW-2026-0001`

如果你想更容易追踪，也可以加入客户简称：

- `LIC-ABC-TRIAL-2026-0001`
- `LIC-ABC-COMM-2026-0002`

---

## 9. 建议维护的内部台账

建议用 Excel 或内部表维护这些字段：

- `license_id`
- `customer_name`
- `customer_id`
- `license_type`
- `binding_mode`
- `issue_date`
- `expiry_date`
- `machine_label`
- `machine_fingerprint`
- `note`
- `current_status`

`current_status` 可用：

- `trial_active`
- `commercial_active`
- `expired`
- `replaced`
- `migrated`

---

## 10. 内部发放检查清单

每次发 `license.json` 之前，至少检查：

1. `customer_name` 是否正确
2. `customer_id` 是否正确
3. `license_type` 是否正确
4. `binding_mode` 是否正确
5. `expiry_date` 是否正确
6. 如果是 `machine_locked`，`machine_fingerprint` 是否正确
7. 生成后是否能正常打开 JSON
8. 文件名是否为 `license.json`

---

## 11. 什么时候应该发哪种 license

### 发 `trial / unbound`

适合：

- 首次试用
- 快速演示
- 临时给客户测试

### 发 `commercial / machine_locked`

适合：

- 正式购买
- 长期交付
- 需要控制只能在指定电脑上运行

---

## 12. 一句话 SOP

- 试用版：直接签发 `trial / unbound` 的 `license.json`
- 正式版：客户先回传 `machine_fingerprint.json`，再签发 `machine_locked` 的 `license.json`
- 续期：通常只需要重新签发新的 `license.json`
- 换电脑：必须重新获取机器指纹，再重新签发
