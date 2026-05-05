# New Project Kickoff Prompt

下面这份提示词模板可以在你开始下一个工具项目时直接复制给 Codex，用来确保新项目先继承当前项目的工程经验，再进入设计和开发。

## 1. 通用版

```text
这是一个新的本地工具项目。先不要写代码。

请先阅读并吸收以下文档中的工程经验，再基于这些经验给出新项目的初始化方案：
- [绝对路径1]
- [绝对路径2]
- [绝对路径3]

输出：
1. 可复用的工程原则
2. 新项目建议沿用的技术栈
3. 新项目建议的目录结构
4. GUI / license / 打包 / 路径治理 / 测试策略
5. 哪些地方不应该机械照搬
6. 建议的 milestone 划分
```

## 2. 针对当前项目经验的推荐版

你可以直接把下面这段作为下个项目的开场提示词：

```text
这是一个新工具项目。先不要写代码。

请先阅读并吸收以下文档中的工程经验，再基于这些经验给出新项目的初始化方案：
- C:\Users\super\capacity_optimizer\docs\TECHNICAL_REFERENCE_CN.md
- C:\Users\super\capacity_optimizer\docs\developer_guide.md
- C:\Users\super\capacity_optimizer\docs\runtime_directory_strategy.md
- C:\Users\super\capacity_optimizer\docs\NEXT_PROJECT_REUSE_GUIDE_CN.md

然后告诉我：
1. 哪些经验应该直接复用
2. 哪些地方需要按新项目调整
3. 你建议的新项目初始架构
4. 你建议的 GUI / license / 打包 / workspace 策略
5. 你建议先做哪些 milestone

在完成上述规划前，不要开始写代码。
```

## 3. 如果你想先讨论需求，不马上开发

```text
进入规划模式。先不要改代码。

请先阅读以下文档，并把里面适合复用到新项目的经验提炼出来：
- C:\Users\super\capacity_optimizer\docs\TECHNICAL_REFERENCE_CN.md
- C:\Users\super\capacity_optimizer\docs\NEXT_PROJECT_REUSE_GUIDE_CN.md

然后基于我的新需求，输出：
1. 新项目目标定义
2. 推荐技术路线
3. 建议目录结构
4. GUI 方案
5. license 方案
6. 打包与交付方案
7. 测试策略
8. milestone 计划
```

## 4. 如果你已经确定要继续 Excel-first

```text
这是一个新的 Excel-first 本地优化工具。先不要写代码。

请先阅读以下文档：
- C:\Users\super\capacity_optimizer\docs\TECHNICAL_REFERENCE_CN.md
- C:\Users\super\capacity_optimizer\docs\NEXT_PROJECT_REUSE_GUIDE_CN.md

然后基于“继续 Excel-first、Windows 桌面交付、支持 license、支持 EXE 打包”的前提，给出：
1. 新项目建议沿用的技术栈
2. 模块分层建议
3. GUI 方案建议
4. workspace / 路径治理建议
5. license 机制建议
6. PyInstaller 打包建议
7. 初始目录结构
8. milestone 计划

在输出这些之前，不要直接实现代码。
```

## 5. 使用建议

最推荐的实际做法：

1. 新项目一开始就先发“先不要写代码”
2. 明确要求先阅读这些经验文档
3. 先让 Codex 输出规划、架构和 milestone
4. 规划确认后，再切换到开发模式

这样可以最大程度复用当前项目的经验，同时避免新项目一开始就走偏。
