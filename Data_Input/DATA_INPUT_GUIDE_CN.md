# Data_Input 数据说明书

本文档用于说明 `Data_Input/` 目录下三类输入数据文件的结构、字段含义和相互逻辑关系。

目的有两个：

1. 让客户可以用自己的真实数据替换当前模板数据。
2. 避免因为漏放数据、列放错、主数据覆盖不完整而导致工具校验报错或运行失败。

---

## 1. Data_Input 目录包含什么

当前目录中有三类数据：

### 1.1 需求清单：Planner Load

- `planner1_load.csv`
- `planner2_load.csv`
- `planner3_load.csv`
- `planner4_load.csv`

说明：

- 这 4 份文件结构完全相同。
- 每一行代表某个 `Product` 在某个月份、某个情景下的需求吨位。
- 系统支持 `planner1_load` 到 `planner6_load`，当前模板只放了 4 份。

### 1.2 产能主数据：Master Capacity

- `master_capacity.csv`

说明：

- 每一行代表某个 `Product` 可以在某个 `Resource` 上生产，并给出该资源对该产品的年产能。

### 1.3 路由主数据：Master Routing

- `master_routing.csv`

说明：

- 每一行代表某个 `Product` 或某个 `Product Family` 在某个 `Resource` 上是否可生产，以及属于 `Primary / Alternative / Toller` 哪一类路线。

---

## 2. 三类数据的关系总览

可以先把三类数据理解成下面这条逻辑链：

```text
Planner Load
  -> 告诉系统 “哪个产品、哪个月份、需要多少吨”

Master Capacity
  -> 告诉系统 “这个产品在哪些资源上有产能、产能是多少”

Master Routing
  -> 告诉系统 “这些可生产资源里，哪些是主路线、哪些是备选路线、哪些是外包路线”
```

更具体一点：

```text
Planner Load 里的 Product
    必须能在 Master Capacity 里找到同 Product

Master Routing 里的 Product / Product Family
    必须能和 Planner Load 里的 Product / ProductFamily 对上

Master Routing 里的 Resource
    如果是内部可生产路线（Primary / Alternative）
    必须能在 Master Capacity 里找到同 Product + 同 Resource
```

---

## 3. 工具实际如何使用这三类数据

### 3.1 ModeA

`ModeA` 只使用两类数据：

- `planner*_load.csv`
- `master_capacity.csv`

在 `ModeA` 下：

- 不看 `master_routing.csv`
- 只要某个产品在 `master_capacity.csv` 有产能记录，就可以参与分配

### 3.2 ModeB

`ModeB` 使用全部三类数据：

- `planner*_load.csv`
- `master_capacity.csv`
- `master_routing.csv`

在 `ModeB` 下：

- 系统会先看路由数据
- `Primary` 路线优先
- `Alternative` 路线次之
- `Toller` 用于外包逻辑

如果某个产品或产品家族在 `master_routing.csv` 里有匹配记录，那么系统会按路由约束来算，不再简单地把所有 capacity 资源都视为可用。

---

## 4. Planner Load 说明

当前模板表头如下：

```text
Month,PlannerName,Product,ProductFamily,Plant,Forecast_Tons,Resource,Scenario Version,Comment
```

### 4.1 每列的含义

#### `Month`

含义：

- 需求所属月份

当前程序接受的常见格式：

- Excel 月份序列值，例如 `46388`
- `YYYY-MM`，例如 `2027-01`
- `YYYY/MM`

建议：

- 客户替换数据时，最好全表统一格式
- 如果来自 Excel 导出，保留当前这种月份序列值也可以

#### `PlannerName`

含义：

- 需求所属 planner / 业务责任人

用途：

- 追溯字段
- 输出展示字段

说明：

- 当前版本要求此列有值
- 但它不参与 capacity 和 routing 的匹配

#### `Product`

含义：

- 产品编码

用途：

- 这是最核心的匹配键之一
- 会和 `master_capacity.csv` 中的 `Product` 做匹配
- 在 `ModeB` 下也可能和 `master_routing.csv` 中的 `Product` 做匹配

注意：

- 同一个产品编码必须在三类文件里保持一致
- 纯数字产品编码如果带前导零，程序会自动去掉前导零
- 最稳妥做法仍然是三类文件统一写法，不要混用

#### `ProductFamily`

含义：

- 产品家族

用途：

- 在 `ModeB` 下，用于和 `master_routing.csv` 的 `Product Family` 做家族级路由匹配

注意：

- 如果客户想按产品家族配置路由，这一列必须写准确
- `ProductFamily` 拼写和内容要与 `master_routing.csv` 的 `Product Family` 一致

#### `Plant`

含义：

- 工厂 / 站点

用途：

- 输出展示和追溯

说明：

- 当前版本要求此列有值
- 但它不参与 capacity / routing 匹配

#### `Forecast_Tons`

含义：

- 该产品该月份的需求吨位

用途：

- 这是需求计算的核心值

规则：

- 必须是数字
- 不能为负数
- 为 `0` 时系统会给 warning

#### `Resource`

含义：

- 需求来源里带出的资源/责任资源说明

非常重要：

- **当前程序不会用 `Planner Load` 里的 `Resource` 去匹配 `master_capacity.csv`**
- 它在当前版本里主要是辅助说明/展示字段

这意味着：

- `Planner Load` 里出现某个 `Resource`
- 并不要求这个资源名字一定出现在 `master_capacity.csv`

客户最容易误解的点就在这里：

```text
Planner Load 的 Resource
!=
Master Capacity / Master Routing 的硬匹配键
```

真正驱动计算的是：

- `Product`
- `ProductFamily`
- `master_capacity.csv`
- `master_routing.csv`

#### `Scenario Version`

含义：

- 情景名，例如基准情景、扩产情景、保守情景

用途：

- 界面中可按情景筛选运行

建议：

- 所有 planner 文件里的情景命名保持一致
- 例如全部使用 `Baseline / Expansion / Lean`

#### `Comment`

含义：

- 备注字段

用途：

- 展示/追溯

说明：

- 不参与计算匹配

### 4.2 Planner Load 必须满足的规则

- 文件名必须符合：`planner1_load.csv` 到 `planner6_load.csv`
- 必须至少有 `Month / Product / Forecast_Tons`
- 实际建议完整保留模板中的所有列
- `Month / Product / PlannerName / Plant` 不要留空
- `Forecast_Tons` 不能是负数

### 4.3 关于重复行

如果同一个文件里出现多条完全相同键值的记录：

- 相同 `Month`
- 相同 `Product`
- 相同 `PlannerName`
- 相同 `Scenario`

系统会自动把吨位合并。

因此建议：

- 客户替换数据前尽量先去重
- 不要把同一个需求拆成多条完全重复的记录

---

## 5. Master Capacity 说明

当前模板表头如下：

```text
Product,Product Family,Resource,Annual Capacity Tons,Utilization Target
```

### 5.1 每列的含义

#### `Product`

含义：

- 产品编码

用途：

- 这是与 `Planner Load` 做产能覆盖匹配的主键

规则：

- `planner*_load.csv` 中出现的每个产品，都应该能在这里找到至少一条对应记录

#### `Product Family`

含义：

- 产品家族

说明：

- 当前程序在 capacity 计算中不依赖这一列做匹配
- 但建议保持正确，便于人工核对

#### `Resource`

含义：

- 生产资源 / 产线 / 工位

用途：

- 在程序内部会被当作 `WorkCenter`
- 与 `master_routing.csv` 中的 `Resource` 一起构成路由和产能的对应关系

规则：

- 如果某个资源会在 `master_routing.csv` 中作为可用路线出现，那么它必须在这里以相同名称出现

#### `Annual Capacity Tons`

含义：

- 该产品在该资源上的年产能

规则：

- 必须大于 `0`
- 单位为吨/年

#### `Utilization Target`

含义：

- 产能利用率目标

规则：

- 可写 `0.88`
- 也可写 `88`，程序会自动识别成 `0.88`
- 必须在 `(0,1]` 范围内，或者可转换到这个范围

### 5.2 系统实际如何使用 Capacity

系统计算时使用的月有效产能为：

```text
Effective Monthly Capacity
= Annual Capacity Tons / 12 * Utilization Target
```

### 5.3 Master Capacity 必须满足的规则

- 每个 `Product` 至少要有一条 capacity 记录
- 同一个 `Product + Resource` 组合不要重复
- `Annual Capacity Tons` 必须大于 0
- `Utilization Target` 必须有效

### 5.4 客户最容易出错的点

错误示例：

```text
Planner Load 中有产品 ORB-01
但 master_capacity.csv 中完全没有 ORB-01
```

结果：

- 系统会报 `NoCoverageCapacity`

---

## 6. Master Routing 说明

当前模板表头如下：

```text
Product,Product Family,Resource,Capacity Ton,EligibleFalg,Router Type
```

注意：

- 模板中的 `EligibleFalg` 是历史拼写，虽然看起来像拼写错误，但程序已经兼容
- 客户替换数据时，建议保留当前表头不改名，最稳妥

### 6.1 每列的含义

#### `Product`

含义：

- 产品级路由

说明：

- 如果这一列有值，表示这条路由规则只针对这个具体产品

#### `Product Family`

含义：

- 产品家族级路由

说明：

- 如果 `Product` 为空，而 `Product Family` 有值，表示该条规则适用于整个产品家族

匹配优先级：

- 产品级路由优先于家族级路由

#### `Resource`

含义：

- 资源 / 产线 / 工位

用途：

- 这是路由实际指向的生产资源

规则：

- 如果这条资源是 `Primary` 或 `Alternative`
- 那么它必须在 `master_capacity.csv` 里存在同一 `Product + Resource` 的 capacity 记录

#### `Capacity Ton`

含义：

- 资源能力说明字段

重要说明：

- **当前版本的求解器不直接使用这一列做容量计算**
- 真正参与容量计算的是 `master_capacity.csv` 中的 `Annual Capacity Tons`

建议：

- 虽然当前不直接参与计算，但仍建议与 `master_capacity.csv` 保持逻辑一致
- 便于人工核查和后续维护

#### `EligibleFalg`

含义：

- 该路线是否可用

当前程序可接受的写法：

- `Y / N`
- `TRUE / FALSE`
- `1 / 0`
- 正数 / 0

当前模板里用的是：

```text
0.88
```

程序会把正数识别为可用。

建议：

- 客户自己维护时，最好改用 `Y` / `N`
- 可读性更高，也更不容易误解

#### `Router Type`

含义：

- 路由类型

可用值：

- `Primary`
- `Alternative`
- `Toller`

业务含义：

- `Primary`：主路线，优先分配
- `Alternative`：备选路线，主路线不够时再用
- `Toller`：外包路线

### 6.2 路由数据如何生效

在 `ModeB` 下：

- 如果某个产品或产品家族在 `master_routing.csv` 中没有任何匹配路由
  - 系统会回退到 capacity-only 逻辑
- 如果有匹配路由
  - 系统就按路由走
  - 不是所有 capacity 资源都自动可用

### 6.3 哪种路由会导致校验错误

#### 情况 A：有路由，但没有任何可用内部路线

例如：

- 某产品在 routing 里有记录
- 但所有 `Primary / Alternative` 都被标为不可用
- 同时也没有产品级 `Toller`

结果：

- 报 `NoCoverageRouting`

#### 情况 B：路由资源在 capacity 里找不到

例如：

- `master_routing.csv` 写了：
  - `Product = AUF-01`
  - `Resource = North Reactor A-1200L`
  - `Router Type = Primary`
- 但 `master_capacity.csv` 中没有：
  - `Product = AUF-01`
  - `Resource = North Reactor A-1200L`

结果：

- 报 `RoutingCapacityMismatch`

这正是最需要提醒客户的地方：

```text
真正必须一一对应的是：
master_routing 的 Resource
和
master_capacity 的 Product + Resource
```

而不是：

```text
planner_load 的 Resource
和
master_capacity 的 Resource
```

### 6.4 只做外包产品时怎么配

如果某个产品不打算走内部资源，只想作为外包产品处理：

- 必须在 `master_routing.csv` 中给这个具体 `Product` 配一条 `Toller`

注意：

- 当前校验里，只有**产品级** `Toller` 会被识别为“允许无内部 capacity”
- 仅写产品家族级 `Toller`，不等价于产品级外包许可

---

## 7. 三类数据之间的硬性对应关系

这是客户替换数据时最重要的一节。

### 7.1 Planner Load -> Master Capacity

规则：

- `planner*_load.csv` 中出现的每一个 `Product`
- 都必须在 `master_capacity.csv` 中至少出现一次

否则：

- 报 `NoCoverageCapacity`

### 7.2 Planner Load -> Master Routing

规则：

- 如果客户要跑 `ModeB`
- 那么建议所有产品都能在 `master_routing.csv` 里找到对应规则
- 匹配方式分两层：
  - 先按 `Product`
  - 再按 `ProductFamily`

注意：

- 如果一个产品/家族在 routing 中完全没有匹配记录，程序会回退到 capacity-only
- 这不一定报错，但可能和客户预期不一致

### 7.3 Master Routing -> Master Capacity

这是最容易漏的关系。

规则：

- `master_routing.csv` 中所有作为 `Primary` 或 `Alternative` 的可用路线
- 都必须在 `master_capacity.csv` 中找到同样的：
  - `Product`
  - `Resource`

否则：

- 报 `RoutingCapacityMismatch`

### 7.4 ProductFamily 一致性

规则：

- 如果 routing 使用家族级规则
- 那么 `planner*_load.csv` 中的 `ProductFamily`
- 必须与 `master_routing.csv` 中的 `Product Family` 完全一致

否则：

- 产品可能找不到家族级路由
- 最终退回 capacity-only 或产生覆盖不完整

---

## 8. 替换客户数据时的推荐步骤

### 第 1 步：先替换 Planner Load

把客户真实需求放进：

- `planner1_load.csv`
- `planner2_load.csv`
- `planner3_load.csv`
- `planner4_load.csv`

替换时先确认：

- 表头不要改
- `Product`、`ProductFamily`、`Plant`、`Forecast_Tons` 都有值
- 月份格式统一

### 第 2 步：生成或整理 Master Capacity

对所有在 planner 中出现的 `Product`：

- 确认每个产品至少有一条 capacity
- 如果一个产品可在多个资源生产，就放多行

### 第 3 步：生成或整理 Master Routing

对所有需要在 `ModeB` 受路由控制的产品：

- 配置 `Primary`
- 必要时配置 `Alternative`
- 需要外包时配置产品级 `Toller`

### 第 4 步：重点核对三张映射表

核对 1：

- `Planner Product` 是否全部出现在 `Master Capacity Product`

核对 2：

- `Routing Product/Product Family` 是否能匹配到 `Planner Product/ProductFamily`

核对 3：

- `Routing Resource` 是否都能在相同产品的 `Master Capacity Resource` 中找到

---

## 9. 建议客户在导入前做的检查清单

正式替换前，建议逐条自查：

- 是否保留了原始表头
- 是否保留了原始文件名
- 是否每个 planner 文件结构一致
- 是否所有 `Product` 都有 capacity
- 是否所有 `Primary / Alternative` 资源都有对应 capacity
- 是否 `ProductFamily` 命名一致
- 是否 `Forecast_Tons` 没有负数
- 是否 `Annual Capacity Tons` 都大于 0
- 是否 `Utilization Target` 合法
- 是否 `Router Type` 只使用 `Primary / Alternative / Toller`

---

## 10. 一个最常见误区

误区：

```text
Planner Load 里的 Resource
必须出现在 Master Capacity 里
```

当前版本中，这个理解**不准确**。

正确理解是：

```text
Planner Load 里的 Resource
主要是说明字段

真正需要严格对应的是：
1. Planner Load 的 Product
   -> Master Capacity 的 Product

2. Master Routing 的 Product / ProductFamily / Resource
   -> Master Capacity 的 Product + Resource
```

---

## 11. 建议客户优先遵循的填报原则

如果客户不想研究所有程序细节，只要记住下面 5 条，通常就不会出大问题：

1. `Planner Load` 负责放需求，核心列是 `Month / Product / ProductFamily / Plant / Forecast_Tons`
2. `Master Capacity` 负责放“产品在哪些资源上有多少产能”，核心列是 `Product / Resource / Annual Capacity Tons / Utilization Target`
3. `Master Routing` 负责放“产品优先走哪些资源”，核心列是 `Product 或 Product Family / Resource / Router Type / EligibleFalg`
4. 每个 planner 里的产品，必须在 capacity 里找到
5. 每个 routing 里的内部资源，必须在 capacity 里找到同产品同资源

---

## 12. 文件命名与运行建议

建议客户保持当前文件名不变：

- `planner1_load.csv`
- `planner2_load.csv`
- `planner3_load.csv`
- `planner4_load.csv`
- `master_capacity.csv`
- `master_routing.csv`

说明：

- 当前程序在 `ModeB` 下会优先找 `alternative_routing`
- 如果没有，再回退找 `master_routing`
- 当前模板直接使用 `master_routing.csv`

因此：

- 客户最稳妥的做法，是继续沿用当前文件名和表头

---

## 13. 最后建议

客户第一次替换真实数据时，建议按下面顺序操作：

1. 先只替换 `planner*_load.csv`
2. 再补齐 `master_capacity.csv`
3. 再补齐 `master_routing.csv`
4. 先跑一次 `ModeA`，确认 capacity 覆盖没问题
5. 再跑 `ModeB`，确认 routing 和 capacity 对应没问题

如果 `ModeA` 能过但 `ModeB` 失败，通常问题就在：

- `ProductFamily` 对不上
- `Router Type` 配置有误
- `master_routing.csv` 里的 `Resource` 在 `master_capacity.csv` 里找不到对应
