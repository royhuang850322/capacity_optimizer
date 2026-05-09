# ModeB 全局 Routing 优化设计说明

## 1. 背景

当前 ModeB 的求解逻辑是分阶段的：

```text
Stage 1：先按 planner load 中的来源资源做基础产能分配
Stage 2：只对 Stage 1 分不掉的 residual 再尝试 routing reroute
Stage 3：仍然分不掉的 residual，再按 Toller / Unmet 分类
```

这个逻辑容易解释，但会错过一种更优的业务解：

```text
产品 A 只能走 EV2
产品 B 可以走 EV2，也可以走 SA3
EV2 不够，但 SA3 有余量
```

当前逻辑可能让产品 B 先占用 EV2，导致产品 A unmet。更理想的结果应是：

```text
产品 B 让出部分 EV2，转去 SA3
产品 A 留在 EV2
最终尽量让两个产品都满足
```

因此，ModeB 后续应从“先占坑，再补救”的分阶段逻辑，升级为“全局 routing 优化”。

## 2. 设计目标

全局 routing 优化的目标是：

```text
把 Primary / Alternative / Toller / Unmet 同时放进同一个优化模型，
由 solver 一次性决定每个产品、每个月、每个工作中心的最优分配。
```

核心业务原则：

1. 尽量满足全部需求。
2. 如果产能紧张，优先把受限资源留给没有替代路径的产品。
3. 有 Alternative 的产品可以被迁移到 Alternative，以释放 Primary 资源。
4. Toller 只在内部 Primary / Alternative 无法满足时使用。
5. Unmet 是最后选择，必须带最高惩罚。

## 3. 目标函数优先级

建议使用分层惩罚的线性目标函数。

优先级从高到低：

```text
1. 最小化 Unmet
2. 最小化 Outsourced / Toller
3. 尽量使用 Primary
4. 其次使用 Alternative
5. 同级 routing 下，按 PenaltyWeight / Priority 选择更优路径
```

对应惩罚建议：

```text
Unmet penalty       = 1,000,000
Toller penalty      = 100,000
Alternative penalty = 1,000 + route penalty
Primary penalty     = 1 + route penalty
```

实际数值可以调整，但必须保证数量级关系稳定：

```text
Unmet >> Toller >> Alternative >> Primary
```

这样 solver 会自然倾向于：

```text
先避免 unmet
再避免外协
再减少 alternative
最后才比较同等级路线的优先级
```

## 4. 决策变量

按月独立求解，每个月建立一组变量。

### 4.1 内部分配变量

```text
x[demand_node, work_center] >= 0
```

含义：

```text
某个月、某产品、某工厂、某来源资源的需求，有多少吨分配到某个内部工作中心。
```

可用的 `work_center` 来自 routing：

```text
RouteType = Primary
RouteType = Alternative
EligibleFlag = Y
```

### 4.2 外协变量

```text
outsource[demand_node] >= 0
```

含义：

```text
某个需求节点有多少吨进入 Toller / Outsourced。
```

只有当该产品存在可用的 `RouteType = Toller` 且 `EligibleFlag = Y` 时，才允许该变量大于 0。

### 4.3 未满足变量

```text
unmet[demand_node] >= 0
```

含义：

```text
某个需求节点最终无法被内部或外协满足的吨数。
```

## 5. 约束规则

### 5.1 需求平衡约束

每个需求节点必须满足：

```text
sum(x[demand_node, work_center])
+ outsource[demand_node]
+ unmet[demand_node]
= demand_tons[demand_node]
```

这保证每一吨需求都有明确归属：

```text
内部生产 / 外协 / 未满足
```

### 5.2 工作中心产能约束

每个月、每个工作中心必须满足：

```text
sum(x[demand_node, work_center] / monthly_capacity[product, work_center]) <= 1
```

这里的关键点是：产能不是简单吨数相加，而是折算为工作中心占用比例。

例如：

```text
产品 A 在 EV2 月产能 = 400 吨
产品 B 在 EV2 月产能 = 100 吨
```

那么：

```text
产品 A 的 1 吨占用 EV2 = 1 / 400
产品 B 的 1 吨占用 EV2 = 1 / 100
```

产品 B 的单位吨位消耗更多 EV2 产能。

### 5.3 Routing 可行性约束

只有 routing 中允许的路径才能建变量。

允许条件：

```text
EligibleFlag = Y
RouteType in {Primary, Alternative}
产品与 WorkCenter 有有效 capacity
```

不允许条件：

```text
EligibleFlag != Y
RouteType = Toller 作为内部分配路径
没有对应 product / work_center capacity
```

### 5.4 Toller 约束

Toller 不进入内部工作中心产能约束。

如果产品存在：

```text
RouteType = Toller
EligibleFlag = Y
```

则允许：

```text
outsource[demand_node] >= 0
```

否则：

```text
outsource[demand_node] = 0
```

## 6. Routing 选择规则

### 6.1 Primary

Primary 是优先路线，但不再是“先占用”的硬规则。

在全局模型中，Primary 是低惩罚路径：

```text
Primary penalty < Alternative penalty
```

因此，只要不会造成更高层级损失，solver 会优先使用 Primary。

### 6.2 Alternative

Alternative 是内部备选路线。

当 Primary 资源紧张时，solver 可以主动把有 Alternative 的产品迁移到 Alternative，以释放 Primary 产能给更受限的产品。

这正是全局 routing 优化要解决的问题。

### 6.3 Toller

Toller 是外协路线。

它不应该与 Primary / Alternative 平级竞争，而应作为内部产能不足之后的次优选择。

### 6.4 Unmet

Unmet 是最后选择。

只要内部和外协路径能满足需求，solver 就不应该产生 unmet。

## 7. 与当前逻辑的差异

### 当前 ModeB

```text
1. 先按来源资源分配
2. 只有剩余 residual 才尝试 Alternative
3. 不会把已经分出去的 Primary 产能拿回来重新安排
```

问题：

```text
有 Alternative 的产品可能占住 Primary，
导致没有 Alternative 的产品 unmet。
```

### 新 ModeB

```text
1. 所有可行路径同时进入模型
2. solver 一次性决定 Primary / Alternative / Toller / Unmet
3. 有 Alternative 的产品可以主动让出 Primary
4. 无 Alternative 的产品优先获得受限 Primary 资源
```

优势：

```text
更接近全局最优
更符合“优先满足全部需求”的业务目标
更少出现可避免的 unmet
```

## 8. 所用工具

### 8.1 Python

继续使用 Python 作为业务编排和数据处理语言。

相关模块：

```text
app/data_loader.py
app/optimizer.py
app/output_writer.py
app/validator.py
```

### 8.2 pandas

用途：

```text
读取 CSV / Excel
标准化输入字段
聚合需求
整理输出明细
生成报表前的数据表
```

### 8.3 OR-Tools Linear Solver

建议继续使用：

```text
ortools.linear_solver.pywraplp
Solver.CreateSolver("GLOP")
```

用途：

```text
建立线性规划模型
定义分配变量
定义产能约束
定义需求平衡约束
最小化 penalty-based objective
```

当前模型仍是连续吨数分配，不需要整数变量，因此 GLOP 仍然适用。

如果未来引入最小批量、整批生产、开线 yes/no 等规则，再考虑切换到：

```text
CBC / SCIP / CP-SAT
```

### 8.4 openpyxl

用途：

```text
输出 Excel 明细
输出 Dashboard
输出月度趋势
输出瓶颈分析
输出产品风险
输出 routing 使用情况
```

## 9. 输出字段建议

为了让用户理解全局 routing 的结果，建议分配明细继续保留并强化这些字段：

```text
Allocation_Source
RouteType
Priority
PenaltyWeight
Source_Resource
WorkCenter
Residual_Unmet_Tons
Capacity_Share_Pct
```

建议新增或强化：

```text
Routing_Decision_Reason
Primary_WorkCenter
Selected_WorkCenter
Primary_Available_Flag
Alternative_Used_Flag
Toller_Used_Flag
```

示例解释：

```text
Selected Alternative because EV2 was constrained and SA3 had available capacity.
```

中文可显示为：

```text
因 EV2 产能受限且 SA3 有可用产能，系统选择 Alternative 路径。
```

## 10. 测试场景

后续落代码时，应至少覆盖以下测试。

### 10.1 无替代产品优先获得受限资源

输入：

```text
产品 A：只能走 EV2
产品 B：可走 EV2 / SA3
EV2 不够，SA3 有余量
```

期望：

```text
A 分配到 EV2
B 部分或全部分配到 SA3
Unmet = 0
```

### 10.2 Alternative 仍不足时才 Unmet

输入：

```text
EV2 不够
SA3 也不够
无 Toller
```

期望：

```text
先用 EV2 和 SA3
剩余才进入 Unmet
```

### 10.3 Toller 在内部不足后使用

输入：

```text
Primary / Alternative 不足
产品有 Toller
```

期望：

```text
剩余进入 Outsourced
Unmet = 0
```

### 10.4 没有 Toller 时才 Unmet

输入：

```text
Primary / Alternative 不足
产品无 Toller
```

期望：

```text
剩余进入 Unmet
```

### 10.5 Primary 优先于 Alternative

输入：

```text
Primary 和 Alternative 都有足够产能
```

期望：

```text
优先使用 Primary
不无故使用 Alternative
```

## 11. 实施注意事项

1. 保留 ModeA 逻辑不变。
2. ModeB 可以新增一条全局求解路径，避免直接破坏旧逻辑。
3. 输出报表需要能解释 Alternative 为什么被使用。
4. 原有 `residual_after_capacity_tons` 和 `residual_after_routing_tons` 字段需要重新定义或标记为 legacy。
5. 新逻辑上线前，应使用当前真实数据回归比较：

```text
旧 ModeB 结果
新 ModeB 全局 routing 结果
Unmet 是否下降
Alternative 是否合理增加
Toller 是否只在必要时出现
```

## 12. 一句话总结

全局 routing 优化的本质是：

```text
不再让产品先抢占 Primary 资源，
而是让 solver 同时看到所有 Primary / Alternative / Toller / Unmet 选择，
从全局角度最小化未满足需求，并在必要时让有替代路径的产品让出受限资源。
```
