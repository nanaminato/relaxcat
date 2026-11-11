# Excel 样式优化指南

适用于本分支的桌面应用和 `tools/RelaxCat.ExcelTool`。本指南以可操作流程、数据边界和验收标准描述能力，不将自动化测试视为生产兼容性认证。发布包的功能以其对应提交为准。

## 支持范围

- 输入：未加密的 `.xlsx` / `.xlsm` OOXML 工作簿；不支持二进制 `.xls`。
- 分析、优化和批量处理均在本地执行，不需要 Excel，也不自动上传工作簿。
- 源路径与输出路径必须不同，扩展名必须一致。非目标包部件按原字节复制，工作表仅改写计划中的样式索引。
- 保留 VBA 二进制不等于验证了宏行为。默认模式遇到 VBA 或未覆盖的包依赖会阻止索引变化；有损模式的明确例外见下文。
- 这是样式表处理工具，不会自动删除图片、合并单元格、条件格式或业务数据，也不保证所有工作簿都能缩小。

## 建议流程

1. 保留原件，在独立输出目录处理副本。关闭可能占用输出文件的应用。
2. 用将要实际执行的模式进行只读分析，检查各分区、预计输出数量、警告及跳过原因。
3. 先优化一份文件，检查输出后的实际指标。单文件输出若已存在会被替换；CLI 不会再次询问，请选新文件名。
4. 在目标 Microsoft Excel 版本中打开、保存并重新打开输出，检查修复提示、外观、公式、打印与依赖该文件的流程。
5. 单文件验收通过后再批量处理。保留原件和独立结果目录，按文件检查失败、跳过和警告记录。

优化先写入目标目录中的 `*.partial` 临时文件，校验通过后移动到输出路径。失败或取消时尝试清理临时文件；若文件被锁定或删除被系统拒绝，可能残留临时文件。工具不提供自动备份或回滚。

## 模式与风险

### 默认优化

默认启用已知引用图内的组件与格式去重，GC 和名称删除均关闭。合并依据是实现所覆盖的等价规则及引用身份，不是预览图看起来相同：

- 字体、填充、边框和自定义数字格式先合并等价定义，再更新相关引用。
- 基础 XF 与单元格 XF 的合并保留命名样式所属关系。不同保留名称不能共享同一个基础 XF，关联不同名称的单元格格式不能仅因当前外观相同而合并。
- 保留默认/保留槽及未引用的唯一项。诊断中的语义重复数可以大于实际可删除数。
- 缺失与显式为 false 的 `apply*` 标志不会混为一谈；有限的等价规范化不代表能证明所有 Excel 行为。
- 未知扩展、未覆盖部件或引用使索引改写跳过，而不是退回不完整的去重。跳过原因通过 `AdvancedOptimizationSkipReason` 返回。

### 未引用样式回收（`--gc`）

默认关闭。保留单元格、行、列的显式/有效样式引用与 `cellXfs[0]`，沿基础 XF 和组件引用回收其余不可达定义。所有保留命名样式仍是根；`cellStyleXfs[0]`、`fonts[0]`、`borders[0]`、`fills[0..1]` 等保留项受保护。

差分格式、表格样式和索引颜色不作为可随意清理的全局组件。差分格式包含内联自定义数字格式时，全局 `numFmts` 保留。无法证明引用覆盖时，不将“未引用”解释为“可以删除”。

### 复制名称清理（`--remove-copied-named-styles`）

默认关闭，独立于 GC。仅考虑具有既存无后缀自定义名称、格式等价、没有额外元数据且未被保留 cell XF 引用的 `Name_正整数` 候选。内置、隐藏、带元数据及有歧义的名称保留。启用 GC 时先处理不可达 cell XF，再判断候选。

先运行 `analyze --json --remove-copied-named-styles`，审核 `NamedStyleCleanupCandidates`。**名称模式不能证明外部程序或模板没有引用该名称。** 该选项可能破坏外部名称依赖，即使 API 的 `IsLossy` 字段仍为 false；目前该字段专门标识破坏性模式。桌面单文件和批量页面也提供默认关闭的选项。

### 破坏性样式清理（`--destructive-styles`）

默认关闭的有损实验功能。保留一个 Normal 身份，将 cell XF 绑定到它并显式启用直接格式，删除其他命名样式和基础 XF，强制执行组件合并与未引用定义回收，无需再加 `--gc`。

它不尝试完整重建被继承的格式，可能改变外观、保护、打印、宏、模板和后续编辑行为。普通 VBA 部件可作为明确行为风险保留；部分 Office 元数据按 [限定结构](destructive-compatibility-validation.md) 识别并保留，存在注音字体引用时整张字体表保持不变。未知依赖与签名等结构性限制仍可阻止优化。详见 [有损模式算法与验证](destructive-style-cleanup.md)。

## 分析指标

- 整个工作簿字节数、`styles.xml` 实际未压缩字节数、ZIP 条目压缩字节数分别报告。
- 分区的 `SerializedBytes` 是解析后独立 UTF-8 序列化大小，可能包含独立命名空间声明，不是源 XML 的精确字节区间。
- `ProjectedCellXfCount`、`ProjectedCellStyleXfCount` 和 `TableChanges` 描述实际计划；唯一数、重复数、未引用数是诊断，不能直接当作节省量。
- `TableChanges` 区分重复删除、未引用回收和有损删除。预计 ZIP 大小是估算，最终以输出实际大小为准。
- 工作表报告包含名称、隐藏状态、范围、单元格计数与样式使用计数。保留样式的空白单元格也会影响这些指标。
- 样式预览用于识别格式，不是 Excel 渲染器。主题色、渐变、图案填充及字体渲染可能与 Excel 不同。

分析和校验逐工作表流式扫描，进度包含工作表与已扫描行数。优化重写会载入单个工作表 XML，超大单表仍可能占用较多内存。性能受样式数量、XML 大小、硬件和启用选项影响；历史人工样本耗时不能用于承诺任意文件的速度。

## 输出校验的含义

输出在提交到目标路径前检查计划一致性、已知引用、工作表结构、未修改部件和输出包：

| 检查键 | 检查内容 |
| --- | --- |
| `stylesSections` | 样式分区与计划一致 |
| `cellXfs` | cell XF 内容与声明数量符合计划 |
| `references` | 工作表样式引用、计数、范围和合并区域符合对应校验规则 |
| `parts` | 未修改部件存在且长度和 SHA-256 一致 |
| `package` | 输出文件存在且非空 |

默认模式的 `VerificationScope=SemanticEquivalenceAndPackageIntegrity` 指已实现模型覆盖的格式等价与包完整性。它不检查外部模板依赖，也不运行 Microsoft Excel。

有损模式的 `IsLossy=true`，范围为 `KnownStructureAndContentOnly_NotVisualOrSemanticEquivalence`。此时 `Verified=true` **不表示**原格式语义或外观保持不变。调用者应同时检查 `Warnings`、`VerificationScope`、`TableChanges` 和跳过原因，不能只看 `Verified`。

## CLI

从仓库根目录运行，使用 `global.json` 指定的 .NET SDK。以下路径为示例：

```powershell
# 分析，不生成优化文件
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- analyze input.xlsx --json

# 默认优化，输出目录必须已存在
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize input.xlsx --output output.xlsx --json

# 在已知引用图内回收
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize input.xlsx --output output-gc.xlsx --gc

# 先审核有损计划，再输出独立文件
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- analyze input.xlsx --destructive-styles --json
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize input.xlsx --output output-lossy.xlsx --destructive-styles --json

# 批量处理；可选 --no-recursive、--overwrite
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- batch input-folder --output-dir output-folder --json

# 只输出有范围限制的扩展类型诊断
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- extensions input.xlsx
```

`analyze`、`optimize`、`batch` 均接受 `--gc`、`--remove-copied-named-styles`、`--destructive-styles`。`--json` 写到标准输出；警告可能写到标准错误，调用方应分别保存。

退出码：`0` 表示命令完成；`1` 表示处理失败、取消或参数解析异常；`2` 用于缺少命令或未知命令。批量退出码不能替代逐文件结果检查，也不代表每个文件都发生了缩减。

## 批量处理

- 默认递归发现 `.xlsx` / `.xlsm`；不跟随重解析点，跳过 `~$` 临时文件和名称以 `-optimized` 结尾的输入。
- 输出目录自动创建，保留相对子目录，文件名追加 `-optimized`。输出位于输入目录内时，扫描排除该输出子树。
- 已有输出默认跳过；只有指定 `--overwrite` 或勾选对应选项才覆盖。输入文件仍保留。
- 单个文件失败会继续处理其他文件。取消时终止后续处理并保留已完成结果；应检查结果中的完成、跳过、失败和取消状态。

## 隐私与问题报告

分析和样式预览不展示单元格值、公式或批注，但完整 JSON、日志与截图可能包含路径、工作表名、命名样式名、格式代码和扩展标识。清空单元格内容不会自动清除这些信息或文档属性。

公开问题报告请优先使用人工生成的最小复现文件。开发和测试仅使用合成数据或用户明确提供并授权使用的样本；不要从“看起来空白”推断可以使用或公开其余内容。

`extensions` 命令只报告有数量/长度限制的类型元数据，不输出普通属性值和正文，但 URI、命名空间、元素名也可能含有敏感标识。该命令不是脱敏保证或安全证明。分享前仍需检查输出，详情见 [扩展诊断范围](excel-style-graph-validation.md#local-extension-type-diagnosis)。

## 故障排查

- **没有明显缩减**：先查看膨胀分区。大量保留命名样式会使基础 XF 继续可达；ZIP 压缩变小不等于样式定义减少。再检查跳过原因。
- **`unsupportedParts` / 扩展或引用跳过**：当前审计没有覆盖相关依赖。不要手动删除 XML 部件来绕过检查；提交经审核的最小复现或类型诊断。
- **输入无效**：索引越界、重复名称/所属基础 XF、分区顺序错误等会被拒绝。不要把失败输出当作修复结果。
- **拒绝访问或文件被占用**：确认输出目录存在且可写，关闭占用输出的应用，检查磁盘空间。尝试普通工作目录；如安全软件拦截，按系统提示审查应用来源和权限，不要关闭整体防护。
- **取消或失败后残留 `.partial`**：确认任务结束且文件不再被使用，再删除残留临时文件；原件不应受影响。
- **Excel 要求修复或外观变化**：停止使用该输出，回到原件。记录工具提交/版本、模式、Excel 版本及修复摘要；审核敏感信息后再提交问题。

## 验收与验证记录

正式使用前，至少检查：Excel 无修复提示；值、公式与关键格式符合预期；合并、行列布局、保护、打印、宏及模板依赖适用时正常；Excel 保存再打开后样式数量与未压缩大小没有异常回涨。

历史记录中的 2026-11-01 Excel COM 样例校验只覆盖当时的单元格 XF 优化，不覆盖后续基础 XF、GC、名称清理或有损模式。当前分支的自动化覆盖和明确未验证项以以下记录为准：

- [样式引用图与命名身份验证](excel-style-graph-validation.md)
- [有损模式的合成基准与验证](destructive-style-cleanup.md)
- [Office 元数据兼容与授权样本验证](destructive-compatibility-validation.md)
- [优化设计与验收标准](excel-style-optimizer-improvements.md)

这些记录不能替代目标 Excel 版本及实际工作流的验收。Windows 发布包打包成功也不等于在 Windows 启动或在 Excel 完成了打开/保存测试。
