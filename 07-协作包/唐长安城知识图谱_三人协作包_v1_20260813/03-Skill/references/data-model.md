# 数据模型速查

## 核心表

- 来源表：`source_key`、题名、作者/机构、年代、资料类型、版本、质量等级、定位方式、审核状态、备注。
- 史料片段：`chunk_key`、`source_key`、章节/页码/图号/行号、原文、现代释义、关联实体、证据角色、审核状态。
- 实体表：`entity_key`、名称、繁体名、类型、别名、类别、时期、描述、位置文本、来源与片段、置信度、知识层、审核状态、备注。
- 关系表：`relation_key`、源实体、目标实体、关系类型、自然语言标签、证据片段、来源、时期、置信度、知识层、审核状态、备注。

## 空间专题层

当前专题结构为四个角度、九张基础专题表，并配一张研究推论表：

1. 地形与水文：水渠表、地形表。
2. 城门与道路：城门表、道路表。
3. 坊市：坊功能表、坊间空间关系表、坊内平面规则表。
4. 重要建筑：重要建筑表。
5. 跨专题：尺度关系表。
6. 研究解释：研究推论表。

专题表仍可增减；新增表前先回答“是否存在一组稳定、重复出现且需要批量审核的专业字段”。若只是少量属性，优先扩展实体属性；若主要表达两对象之间的连接，优先使用关系表。

## 建议实体类型

`Fang`、`Market`、`Road`、`Canal`、`Gate`、`Area`、`Building`、`TerrainFeature`、`WaterBody`、`LayoutRule`、`ScaleComparison`、`FunctionalZoningRule`、`InterpretiveClaim`、`Person`。

## 协作登记对象与正式实体的区别

协作表中的“主实体登记表”是编号、名称和查重的公共目录，不是可直接发布的正式实体表。

- `Fang`、`Road`、`Canal`、`Gate`：可先登记稳定编号和最小识别信息；只有来源、片段与必要字段齐全后，才转入正式实体表。
- `ScaleComparison`：先登记比较编号和比较对象名称；必须补齐两个比较对象、指标、单位、可比性和证据后，才转入正式实体/关系系统。
- `Rule`：先登记规则编号、名称和范围；必须补齐条件、结论、正例、例外、推理过程、来源与片段后，才作为 `FunctionalZoningRule` 或 `InterpretiveClaim` 转入正式系统。
- “待补字段”“待补证据”“待核对”均不等于审核通过，不能进入发布白名单。

## 常用关系类型

- 空间：`LOCATED_IN`、`EAST_OF`、`WEST_OF`、`NORTH_OF`、`SOUTH_OF`、`ADJACENT_TO`、`NEAR`、`FACING`。
- 网络：`CONNECTS_TO`、`FLOWS_THROUGH`、`SUPPLIES_WATER_TO`。
- 组成与归属：`PART_OF`、`HAS_PART`。
- 解释：`INTERPRETS`、`INVOLVES`、`SUPPORTED_BY`、`REFINES`。
- 人事：`ASSOCIATED_WITH`、`INVOLVED`。

关系方向必须由标签可直接读出。例如“平康坊位于崇仁坊南侧”应为平康坊指向崇仁坊的 `SOUTH_OF`。

## 证据角色

- `direct`：原文明示该事实或关系。
- `tabular`：来源表格直接记录。
- `figure_reading`：从可靠复原图读取；必须记录图号和解读方法。
- `archaeological`：考古报告或测量证据。
- `interpretive`：现代研究者的论证或推断。
- `counterevidence`：反例、限制或相反意见。

实体或关系可以引用多个片段；每个片段也可以支持多个条目。不要为了“一条关系一个片段”而重复抄录相同原文。
