# 下一个项目标准提示词

本文档用于把当前项目已经验证过的工程方案，转成可以直接复制给 Codex 的标准提示词。

适用范围：

- 本地 Windows 工具
- Python 后端
- 桌面 GUI
- Excel / CSV 输入输出
- 需要打包成 exe
- 需要明确 install 目录和 workspace 目录边界

## 1. 继承当前方案的标准总提示词

```text
这是一个新的本地 Windows 业务工具项目。先不要写代码。

请按照下面这套已经验证过的工程方案，为新项目给出完整的初始化设计：

技术与交付约束：
- 核心语言使用 Python
- 桌面 UI 使用 PySide6（Qt Widgets），不要用 Web UI
- Excel / CSV 作为主要输入输出载体
- 使用 pandas 做数据整理
- 使用 openpyxl 做 Excel 读写、样式、报表生成
- 如果需要优化求解，使用 OR-Tools
- 打包方案使用 PyInstaller one-folder，不要用 one-file
- 运行路径治理集中在单独模块，不要把路径规则散落在业务代码里
- 区分 install 目录和 user workspace 目录
- user workspace 至少包含：
  - Tooling Control Panel
  - Data_Input
  - output
  - logs
  - licenses
  - docs
- GUI 只负责收集参数、初始化 workspace、触发运行、打开输出目录，不要在 GUI 里放求解或报表逻辑
- 所有日志集中写到 logs 目录
- 桌面发布目标是给 Windows 最终用户直接双击 exe 使用

请基于这些约束输出：
1. 新项目应该直接复用的工程原则
2. 新项目建议沿用的技术栈
3. 建议的目录结构
4. GUI / runtime paths / workspace / logging / packaging 方案
5. 哪些地方应该复用，哪些地方只能借鉴不能照搬
6. 按 milestone 拆分的实施计划

在完成上述规划前，不要直接开始实现代码。
```

## 2. 强制沿用 PySide6 + PyInstaller 的提示词

```text
这是一个新的桌面工具项目。请直接采用以下固定方案，不要改成 Web、Electron、Tkinter 或其他替代路线：

- GUI：PySide6（Qt Widgets）
- 打包：PyInstaller one-folder
- 入口：.pyw 薄入口 + app/...launcher.py 真正 UI 逻辑
- 资源打包：通过单独的 packaging manifest / spec 管理
- 版本信息：通过独立 version 文件写入 exe 元数据

请输出：
1. 为什么这套方案适合本地 Windows 企业工具
2. GUI 层、业务层、报表层、路径层、打包层的边界
3. PySide6 窗口组织方式
4. .pyw 到 exe 的完整打包链路
5. 发布时应交付哪些 zip / exe / 文档
6. 测试和验收清单

不要先写代码，先给架构和交付设计。
```

## 3. 强制沿用 workspace 路径治理的提示词

```text
这是一个新的本地 Windows 工具项目。请严格沿用“集中式 runtime path + user workspace”方案。

要求：
- 路径解析集中在一个模块，例如 runtime_paths.py
- 所有代码都通过统一的 RuntimePaths 对象拿路径
- 不允许在业务逻辑里到处拼相对路径
- source mode 可以使用仓库根目录作为 workspace
- packaged mode 默认使用 user workspace
- workspace 目录至少包含：
  - Tooling Control Panel
  - Data_Input
  - output
  - logs
  - licenses
  - docs
- GUI 启动时可以初始化 workspace，但不能覆盖已有用户文件
- 所有用户可写内容都应进入 workspace，而不是 install 目录

请输出：
1. 路径策略设计
2. RuntimePaths 数据结构建议
3. source mode 与 packaged mode 的差异
4. workspace 初始化规则
5. 需要的 smoke tests
6. 常见错误和防御策略

不要写实现代码，先给路径治理方案。
```

## 4. 继承当前 GUI 边界规则的提示词

```text
请为新项目设计一个 PySide6 桌面 launcher，并严格遵守以下边界：

GUI 负责：
- 选择工作目录
- 选择输入文件 / 报告文件
- 输入参数
- 保存 launcher settings
- 调用后端函数
- 打开 output / logs 目录
- 弹出友好的错误信息

GUI 不负责：
- 优化求解
- Excel 报表计算
- 业务分配逻辑
- 路径规则定义

请输出：
1. 建议的窗口结构
2. 建议的控件分区
3. settings 保存策略
4. 后端调用边界
5. 错误提示与日志策略
6. 哪些逻辑必须留在 app 层，不要塞进 launcher
```

## 5. 继承当前打包与发布流程的提示词

```text
请按本地 Windows 企业工具的发布方式，为新项目设计打包与交付流程。

强制约束：
- 使用 PyInstaller one-folder
- 每个桌面工具都有单独 spec 文件
- 用统一的 PowerShell 构建脚本驱动打包
- 产物输出到 dist/
- 分发压缩包输出到 delivery_packages/
- 版本号集中维护
- exe 写入 Windows version metadata
- 打包后必须有自动校验步骤，确认 one-folder 布局和资源完整性

请输出：
1. packaging 目录结构
2. manifest / spec / version 文件如何组织
3. build_onefolder.ps1 应承担的职责
4. zip 命名规范
5. 发布前验证清单
6. Git tag / release notes 建议流程
```

## 6. 复用当前项目经验，但不要盲抄的提示词

```text
这是一个当前项目的后继工具。请基于已有项目经验，明确区分以下三类内容：

1. 应直接复用
2. 应按新项目适配后复用
3. 不应照搬

当前已知应优先复用的工程方案：
- Python + PySide6 + openpyxl + pandas
- PyInstaller one-folder
- runtime_paths 集中路径治理
- workspace 初始化模型
- GUI 与业务逻辑分层
- logs / licenses / output 的目录治理
- smoke test 与 packaging validation

当前已知不应盲抄的内容：
- 旧文件名
- 旧列名
- 旧业务模式命名
- 旧报表口径
- 旧客户特定字段

请输出：
1. 直接复用项
2. 适配复用项
3. 禁止照搬项
4. 建议的初始目录结构
5. 建议的 milestone 计划
```

## 7. 推荐使用顺序

建议你在下个项目里按这个顺序使用：

1. 先发“标准总提示词”
2. 如果方向已确定，再补“PySide6 + PyInstaller 固定方案提示词”
3. 如果路径和安装策略重要，再补“workspace 路径治理提示词”
4. 开始写 GUI 前，再补“GUI 边界规则提示词”
5. 准备交付前，再补“打包与发布流程提示词”

## 8. 当前项目中可参考的文件

- [README.md](/C:/Users/super/capacity_optimizer/README.md)
- [docs/TECHNICAL_REFERENCE_CN.md](/C:/Users/super/capacity_optimizer/docs/TECHNICAL_REFERENCE_CN.md)
- [docs/developer_guide.md](/C:/Users/super/capacity_optimizer/docs/developer_guide.md)
- [docs/runtime_directory_strategy.md](/C:/Users/super/capacity_optimizer/docs/runtime_directory_strategy.md)
- [docs/desktop_launcher_usage.md](/C:/Users/super/capacity_optimizer/docs/desktop_launcher_usage.md)
- [docs/pyinstaller_onefolder_build.md](/C:/Users/super/capacity_optimizer/docs/pyinstaller_onefolder_build.md)
