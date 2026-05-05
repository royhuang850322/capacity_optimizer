# 项目交接检查清单

本清单用于在把 `capacity_optimizer` 交给下一位开发者前，快速确认交接材料是否齐全。

## 1. 仓库与版本

- [ ] 已确认交接用的 GitHub 仓库地址
- [ ] 已确认要交接的分支/提交/tag
- [ ] 已确认本地未提交改动是否要带上

## 2. 文档

- [ ] 已提供 [PROJECT_HANDOFF_CN.md](/C:/Users/super/capacity_optimizer/docs/PROJECT_HANDOFF_CN.md)
- [ ] 已提供 [TECHNICAL_REFERENCE_CN.md](/C:/Users/super/capacity_optimizer/docs/TECHNICAL_REFERENCE_CN.md)
- [ ] 已提供 [developer_guide.md](/C:/Users/super/capacity_optimizer/docs/developer_guide.md)
- [ ] 已提供 [runtime_directory_strategy.md](/C:/Users/super/capacity_optimizer/docs/runtime_directory_strategy.md)
- [ ] 已提供 [NEW_PROJECT_KICKOFF_PROMPT_CN.md](/C:/Users/super/capacity_optimizer/docs/NEW_PROJECT_KICKOFF_PROMPT_CN.md)
- [ ] 已提供 [NEXT_PROJECT_REUSE_GUIDE_CN.md](/C:/Users/super/capacity_optimizer/docs/NEXT_PROJECT_REUSE_GUIDE_CN.md)

## 3. 运行与打包

- [ ] 已说明如何安装依赖
- [ ] 已说明如何启动 Launcher
- [ ] 已说明如何生成报告
- [ ] 已说明如何重新打包 EXE
- [ ] 已说明哪些目录是运行产物

## 4. license

- [ ] 已说明 `licenses/active/license.json` 的用途
- [ ] 已说明 machine fingerprint 流程
- [ ] 已说明内部签发流程文档位置

## 5. 输入输出样例

- [ ] 已保留一份代表性的 `Data_Input`
- [ ] 已保留一份代表性的 `ModeA` 输出样例
- [ ] 已保留一份代表性的 `ModeB` 输出样例
- [ ] 已保留一份代表性的 `Summary` 输出样例

## 6. 给下一位开发者的 Codex 开场方式

- [ ] 已把建议的 Codex 开场提示词一起发给下一位开发者
- [ ] 已明确要求“先读文档，先规划，再写代码”

## 7. 交接完成标准

满足以下条件才算真正完成交接：

- [ ] 下一位开发者知道从哪份文档开始看
- [ ] 下一位开发者知道如何本地运行项目
- [ ] 下一位开发者知道如何重新打包项目
- [ ] 下一位开发者知道 license 的流转方式
- [ ] 下一位开发者知道如何让 Codex 继续接手开发
