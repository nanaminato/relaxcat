# Excel 样式优化：后续改进计划

> Status: Proposal
>
> 本文仅记录后续改进方向，不修改当前实现。

## 背景与最新定位

当前 Excel 样式优化第一版主要针对 `xl/styles.xml` 中的 `cellXfs` 做完全重复项去重，并同步更新工作表中的样式索引。该方案安全、可验证，但实际工作簿测试表明：`cellXfs` 从 184 降至 179 后，`styles.xml` 的**未压缩大小仍约为 9.35 MB**。

进一步直接检查 OOXML 后发现，大量 `cellXfs/xf` 含有 `xfId="48845"`。OOXML 索引从 0 开始，因此该文件的 `cellStyleXfs` 至少存在 48,846 个条目。相比之下，当前 `cellXfs` 只有约 184 个，因此现有算法只处理了很小的一部分样式数据。

同时确认：优化后 XLSX 文件体积下降主要来自 ZIP 压缩率提高，而不是 `styles.xml` 未压缩内容显著减少。Excel/TestHelper 再次保存后文件体积可能回归，因此**压缩后的 XLSX 大小不能单独作为样式优化是否有效的判断依据**。

基于现有证据，下一阶段应优先调查 `cellStyleXfs`，同时保留对其它样式分区的完整诊断能力。

> 公司测试工作簿与 TestHelper 不进入个人开发环境。真实文件只在公司环境本地分析；开发与自动测试使用脱敏的结构统计和人工生成的 synthetic pathological workbook，不复制业务内容、公式、文件名或路径。

## 1. 增加 styles.xml 分区诊断

分析阶段应分别统计：

- `numFmts`
- `fonts`
- `fills`
- `borders`
- `cellStyleXfs`
- `cellXfs`
- `cellStyles`
- `dxfs`
- 其它样式相关分区

除元素数量外，同时计算：

- XML 未压缩字节大小或可靠的分区序列化大小；
- ZIP 内压缩后大小（仅作为辅助指标，不与未压缩大小混淆）；
- 被引用元素数量；
- 未引用元素数量；
- 完全重复元素数量；
- 各分区占整个 `styles.xml` 的比例。

对于 `cellStyleXfs`，额外统计：

- `cellStyleXfs` 总数与最大有效索引；
- 被 `cellXfs.xfId` 引用的索引集合及引用次数；
- 被 `cellStyles.xfId` 引用的索引集合及引用次数；
- 同时被多个来源引用的项；
- 完全未引用项；
- 完全重复项以及去重后的预计数量/大小。

界面应明确区分：

```text
styles.xml uncompressed     9.35 MB
styles.xml ZIP compressed   509 KB -> 445 KB
cellXfs                     184 -> 179
cellStyleXfs                >= 48,846   <- investigate first
```

避免把 ZIP 重打包带来的压缩率提升误认为样式数据已经被真正删除。

## 2. 优先建立 cellStyleXfs 引用图

下一阶段的首要目标是完整理解 `cellStyleXfs` 的引用关系，而不是直接删除 48,846 个条目。

至少覆盖：

```text
cellStyles.xfId ───────────┐
                           ↓
                     cellStyleXfs
                           ↑
cellXfs.xfId ──────────────┘
      ↑
worksheet c/row/col -> cellXfs
```

还应检查 OOXML 中其它可能引用或依赖命名样式/基础 XF 的结构。只有能够证明引用关系完整时，才能重映射或删除条目。

诊断结果应回答：

1. 48,846 个 `cellStyleXfs` 中实际唯一的定义有多少？
2. 有多少被 `cellXfs.xfId` 实际引用？
3. 有多少被 `cellStyles.xfId` 实际引用？
4. 有多少完全不可达？
5. 最大 `xfId=48845` 是否只是大量历史样式累积后的最后一个基础样式？
6. 删除完全重复/不可达项后，预计 `styles.xml` 能减少多少？

## 3. cellStyleXfs 安全去重

在引用图验证完成后，优先实现 `cellStyleXfs` 的完全重复项去重：

1. 为每个 `cellStyleXfs/xf` 生成规范化结构键；
2. 仅合并规范化 XML 完全一致的定义；
3. 建立 `old xfId -> new xfId` 映射；
4. 更新所有 `cellXfs.xfId`；
5. 更新所有 `cellStyles.xfId`；
6. 重建 `cellStyleXfs` 并更新 `count`；
7. 对所有引用执行越界与结构一致性校验。

第一阶段不要做“视觉上相同”的激进合并，只处理能够证明结构完全等价的项。

## 4. 底层样式组件去重

完成 `cellStyleXfs` 定位后，再根据诊断数据决定是否扩展到：

- `fonts`
- `fills`
- `borders`
- `numFmts`
- 其它组件

处理方式仍为完全等价去重：建立 `oldId -> newId` 映射，更新 `cellXfs` 与 `cellStyleXfs` 中的组件引用，然后再次对 XF 进行去重。

不应预先假定 `borders`、`fonts` 或 `fills` 是主要膨胀来源；以实际分区大小与引用统计为准。

## 5. 安全垃圾回收（GC）

在完整引用图覆盖后，可以删除能够证明不可达的样式定义，包括：

- 未引用 `cellStyleXfs`；
- 未引用 `cellXfs`；
- 未引用 `fonts`；
- 未引用 `fills`；
- 未引用 `borders`；
- 其它确认无隐藏依赖的定义。

无法证明安全的定义必须保留。特别检查命名样式、条件格式、差分格式、表格、数据透视表及其它可能间接依赖样式的 OOXML 部件。

## 6. 推荐实现阶段

### Phase 1：只读诊断（最高优先级）

- 分区 count 与未压缩大小；
- ZIP 压缩大小单独显示；
- `cellStyleXfs` 引用数、未引用数、重复数；
- 明确识别主要膨胀来源；
- 不改变任何优化行为。

### Phase 2：Synthetic pathological workbook

在个人开发环境生成不含公司数据的测试工作簿，例如：

- 数万 `cellStyleXfs`；
- 大量完全重复基础 XF；
- 少量 `cellXfs` 指向高索引 `cellStyleXfs`；
- 混合 `cellStyles.xfId` 与 `cellXfs.xfId` 引用；
- 可配置 fonts/fills/borders 污染。

用于自动验证大规模样式表的分析、映射、去重和性能。

### Phase 3：cellStyleXfs 完全重复项去重

在引用关系测试充分后，执行保守的结构完全等价去重，并更新所有已知 `xfId` 引用。

### Phase 4：安全 GC 与底层组件去重

根据实际诊断结果扩展不可达项清理及 fonts/fills/borders 等组件去重。

### Phase 5：Excel round-trip 稳定性验证

公司环境可使用原有 TestHelper 或 Excel 完成：

```text
A = 原文件
B = RelaxCat 优化文件
C = B 经 Excel/TestHelper 打开并保存后的文件
```

比较 A/B/C 的未压缩 `styles.xml` 大小、各分区 count、引用统计和最终 XLSX 大小。

最终优化结果应尽可能满足：Excel 重新保存后，已经删除的冗余样式不会大规模重新生成。该验证属于最终验收，不要求将 TestHelper 或公司工作簿带入开发环境。

## 7. 校验要求

扩展优化范围后至少验证：

- 所有 `cellXfs.xfId` 均指向有效 `cellStyleXfs`；
- 所有 `cellStyles.xfId` 均指向有效 `cellStyleXfs`；
- XF 中 `fontId` / `fillId` / `borderId` / `numFmtId` 等引用有效；
- 所有工作表样式索引有效；
- 各分区 `count` 与实际子元素数量一致；
- 优化前后实际使用样式结构/外观等价；
- 非目标 OOXML 部件不发生意外变化；
- `.xlsm` 的 VBA 等二进制部件保持原样；
- Excel 能正常打开且无需修复；
- ZIP 压缩率变化不被计入“样式数据删除量”。

## 目标

后续优化的核心指标应改为：

- `styles.xml` **未压缩大小**实际减少量；
- `cellStyleXfs` 总数、唯一数、引用数和不可达数变化；
- `cellXfs` 与底层组件的实际减少量；
- XLSX/XLSM 最终文件大小（辅助指标）；
- Excel round-trip 后优化结果是否稳定；
- 工作簿格式与行为是否保持不变。

当前案例已提供一个明确方向：`cellXfs 184 -> 179` 并不是主要优化空间，而 `xfId=48845` 暗示的至少 48,846 个 `cellStyleXfs` 才是下一阶段首先需要解释和安全处理的对象。
