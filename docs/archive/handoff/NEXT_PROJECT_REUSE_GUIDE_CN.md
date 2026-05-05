# Next Project Reuse Guide

本文档用于指导后续新工具项目如何复用当前 `capacity_optimizer` 项目的工程经验，减少重复试错。

适用场景：

- 新建一个类似的本地优化工具
- 新建一个 Excel-first 的桌面分析工具
- 新建一个需要 license、打包、桌面交付的企业内部工具

## 1. 建议默认复用的工程原则

以下原则建议作为下一个项目的默认起点：

1. 保持 Excel-first，而不是默认改成 Web
2. GUI 只做流程编排，不承载核心业务计算
3. 路径治理必须集中化
4. 只暴露一个主 workspace 路径给用户
5. 打包优先使用 PyInstaller one-folder
6. license 继续采用离线签名方案
7. 测试优先覆盖启动、路径、输出、打包、license

## 2. 哪些技术栈建议直接沿用

如果新项目仍然是“本地 Windows 企业工具”，建议默认沿用：

- Python
- pandas
- openpyxl
- PySide6
- PyInstaller
- cryptography

如果新项目仍涉及优化求解，再继续沿用：

- OR-Tools

## 3. 哪些架构方案建议直接沿用

建议继续使用“模块化单体应用”架构，至少保留这些模块边界：

- `app/main.py`
  CLI 或主入口编排
- `app/desktop_launcher.py`
  GUI 启动器
- `app/runtime_paths.py`
  路径治理中心
- `app/workspace_init.py`
  工作区初始化
- `app/data_loader.py`
  输入数据读取与标准化
- `app/optimizer.py`
  核心求解器
- `app/output_writer.py`
  Excel 报表输出
- `app/license_validator.py`
  license 校验

核心经验：

- 不要把 GUI、求解、报表混到一个文件里
- 不要让业务逻辑散落到 bat、Excel、GUI 事件中

## 4. GUI 方案建议

如果下一个工具仍然是本地 Windows 交付，建议默认使用：

- `PySide6 + Qt Widgets + QSS`

原因：

- 比 Tkinter 更适合企业桌面软件
- 比 Web 方案更适合离线环境
- 更适合后续打包为 EXE

GUI 默认只承担：

- 路径设置
- 参数设置
- 工作区初始化
- 运行触发
- 结果与日志入口

GUI 不要承担：

- 求解逻辑
- 业务核心计算
- Excel 报表生成逻辑

## 5. 路径与目录策略建议

下一个项目建议继续沿用“一个主路径 + 派生目录”的策略。

建议目录模型：

- `workspace`
- `workspace/Data_Input`
- `workspace/output`
- `workspace/logs`
- `workspace/licenses`
- `workspace/docs`

建议继续通过统一模块治理：

- 安装目录
- 资源目录
- workspace
- logs
- output
- licenses
- docs

关键原则：

- 用户写入内容全部进入 workspace
- 安装目录尽量只读
- 不让用户配置多个彼此关联的路径

## 6. License 方案建议

如果新项目仍需要离线交付和权限控制，建议直接沿用当前模型：

1. 客户端生成 machine fingerprint
2. 内部签发 `license.json`
3. 客户放到指定 `licenses/active` 目录
4. 程序启动时进行签名、期限、机器绑定、产品范围校验

适合：

- 试用版
- 正式版
- 限期版本
- 机器绑定版本

经验：

- license 不只是一个文件格式问题，而是一个完整运营流程
- 一定要同时设计 request、issue、active、archive 的管理方式

## 7. 打包与交付策略建议

对 Windows 企业环境，建议继续默认使用：

- `PyInstaller one-folder`

不建议默认使用：

- one-file
- 需要客户先装 Python 的源码交付

推荐交付分层：

- 开发态：源码仓库
- 内部测试态：打包后的 dist 目录
- 客户交付态：one-folder 压缩包

## 8. 测试策略建议

下一个项目不要一开始只写业务代码，建议尽早补最小 smoke tests。

优先覆盖：

- 启动是否成功
- 路径初始化是否成功
- 工作区是否创建成功
- 输入是否能读取
- 输出目录是否能写入
- 打包后资源是否可定位
- license 是否能正确校验

经验：

- 桌面工具的最大风险通常不是算法，而是交付链路
- 所以测试要优先贴近真实使用路径

## 9. 哪些地方不要机械照搬

以下内容不要无脑复制，要根据新项目调整：

1. 输入文件格式
   新项目的数据列名、业务实体、粒度可能不同
2. 报表结构
   不同项目的 dashboard 和 summary 不一定相同
3. 求解模型
   ModeA / ModeB 只是当前项目的业务模式，不应默认照搬
4. UI 字段布局
   新项目的设置项不一定需要当前这么多分区
5. license 字段
   customer id、product code、binding scope 可能需要重新定义

## 10. 推荐的新项目启动顺序

建议下一个项目按这个顺序启动：

1. 先明确输入输出和客户使用路径
2. 再决定是否继续 Excel-first
3. 再确定 GUI、license、打包、workspace 策略
4. 再划分模块边界
5. 最后才进入具体业务逻辑和求解开发

## 11. 推荐你在新项目开场时给 Codex 的要求

建议你在新项目第一轮就明确提出：

- 先阅读当前项目经验文档
- 先做规划，不直接写代码
- 先给新项目的初始架构建议
- 明确哪些经验复用，哪些不复用

建议搭配使用：

- [TECHNICAL_REFERENCE_CN.md](/C:/Users/super/capacity_optimizer/docs/TECHNICAL_REFERENCE_CN.md)
- [NEW_PROJECT_KICKOFF_PROMPT_CN.md](/C:/Users/super/capacity_optimizer/docs/NEW_PROJECT_KICKOFF_PROMPT_CN.md)
