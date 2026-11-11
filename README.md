# RelaxCat

Windows 截图收集与 Excel 排版工具。支持图片或文件夹拖入、文件夹监控、手动导入，提供图片卡片与放大预览、DPI 感知缩放、有序规则、逐行/节省空间排列，以及 Microsoft Excel COM 导出；可选择捆绑的 Greenshot 截图。另提供完全离线的 Excel 样式分析/去重工具，显示工作表样式与优化前后效果，并输出到新工作簿。界面使用 Avalonia XAML + C#，支持明亮/黑暗主题和简体中文、日本語、English、한국어切换。

## 版本与下载

本分支包含默认关闭的**有损实验功能**。文档描述本分支源码，不代表默认分支或任意历史安装包已经包含这些功能；自动化校验通过也不等于完成 Microsoft Excel 兼容性验收。

在 [Releases](https://github.com/nanaminato/relaxcat/releases) 查看实际可用的版本、对应源码和验证范围。预发行版仅用于评估，不将旧版本的验证结果套用到新算法。若没有符合需求的附件，请按下文从源码构建。

Windows x64 便携包采用以下命名；具体提供哪些附件以该版本发布页为准。解压后进入 `RelaxCat-v<版本>-win-x64` 文件夹运行 `RelaxCat.App.exe`：

- **完整便携版**（`RelaxCat-v*-win-x64.zip`）：包含 .NET 10 运行时，无需另行安装。
- **轻量版**（`RelaxCat-v*-win-x64-framework-dependent.zip`）：不含运行时，需要安装 [.NET 10 x64 Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)。不要求 Desktop Runtime；只有 .NET 8/9 不能代替 .NET 10。

需要 Windows 10 或更新版本、x64。图片导出需要本机安装桌面版 Microsoft Excel；图片导入、预览及离线样式分析/优化不需要 Excel。捆绑 Greenshot 另需 .NET Framework 4.8。

## 界面预览

以下历史截图使用合成样例展示布局；具体选项与文案以运行版本为准，截图中的计数不代表其他工作簿的优化效果。

| 图片工作台（浅色） | 图片工作台（深色） |
| :---: | :---: |
| ![图片工作台](docs/images/workspace-light.png) | ![深色主题](docs/images/workspace-dark.png) |

| 放大预览 | Excel 排版设置 |
| :---: | :---: |
| ![放大预览](docs/images/preview.png) | ![Excel 排版设置](docs/images/settings-excel.png) |

| Excel 样式分析 | 优化结果（明亮） | 优化结果（黑暗） |
| :---: | :---: | :---: |
| ![Excel 样式分析](docs/images/excel-style-analysis.png) | ![优化结果](docs/images/excel-style-result.png) | ![深色主题](docs/images/excel-style-result-dark.png) |

| Excel 批量优化 |
| :---: |
| ![批量优化](docs/images/excel-batch.png) |

## 从源码运行

需要构建用 .NET SDK 10.0.300（见 `global.json`，允许同一功能带的补丁更新）。SDK 已包含运行时；桌面应用与 Excel COM 工作流请在 Windows 上运行。

```powershell
dotnet run --project src/RelaxCat.App -c Release
```

也可启动构建生成的 `src/RelaxCat.App/bin/Release/net10.0-windows/RelaxCat.App.exe`。

1. 点击「导入图片」选择本地 PNG、JPEG 或 BMP，或把多张图片、文件夹直接拖入窗口——文件夹会递归导入子文件夹中的图片。首页以卡片展示每张图片；取消「参与导出」勾选可保留图片但不导出。卡片上的「复制」复制原始图片到剪贴板，「移除」只移出队列，不删除源文件。
2. 点击缩略图打开放大预览：支持适应窗口、缩放与滚动、上一张/下一张，左右方向键切换，Esc 或点击背景关闭。预览中的勾选框与卡片同步。
3. 首页提供默认缩放比例、节省空间排列、排序与全部移除。点击左侧「设置」或「详细设置」，在「尺寸与规则」「Excel 排版」「外观与语言」独立设置页面调整其他选项。空模板使用 Excel 默认模板。顶部「切换明暗」也可快速切换主题。
4. 点击「导出到 Excel」，仅按队列顺序导出已勾选图片。首次创建工作簿，后续复用，导出前清空目标工作表中的全部图片并重新排版（保留单元格内容和非图片图形）；如果已关闭则重新创建。导出前会重新规范工作表行高、列宽，可能影响该工作表的原有内容。
5. 在 Excel 中自行保存。关闭 RelaxCat 不会强行关闭已导出的工作簿。主题、语言、导出、截图和排序设置会自动保存，重启后恢复；图片队列仍仅保留在当前会话。连续输入会合并保存，关闭时立即保存。设置文件带版本号和备份，主文件损坏时尝试备份，两者都无效时使用默认设置并提示。

## 工作区整理

- 默认按文件最后修改时间从早到晚排列，导入完成后自动插入到正确位置，也适用于监控收到的新截图。可以切换为从晚到早或导入顺序；相同时间保持导入顺序。文件时间在导入时读取，之后修改磁盘文件不会自动重新排序。
- 也可以把多张图片或文件夹拖入窗口导入：文件夹按名称顺序递归展开子文件夹，同一张图片重复拖入只导入一次，非图片文件忽略；无法读取的文件夹在状态栏提示，其余内容照常导入。
- 导出和预览沿用当前工作区顺序；排序不会改变图片的导出勾选状态。
- 「全部移除」清空所有图片并释放预览和缩略图，只移出工作区，保留源文件及 Excel 中已导出的内容。导入或导出过程中不可清空；监控继续运行，后续新截图仍会进入工作区。

## 截图与文件夹监控

1. 点击左侧「截图与监控」，选择「Greenshot（随应用捆绑）」或「仅监控文件夹」，指定截图文件夹，然后点击「应用并启动」。默认关闭，不会在未选择时接管系统截图。
2. Greenshot 模式首次启动会从随应用分发的 ZIP 校验并解压官方便携预发行包，无需另装 Greenshot。当前捆绑版本为 `1.4.371-g78211b16d9`，适用于 Windows 10+ x64，需要 .NET Framework 4.8。来源、校验和及上游许可证位于 `vendor/greenshot/`，发布目录也包含这些文件。
3. 首次初始化后保留 Greenshot 的独立配置。截图快捷键、输出方式和保存目录请在 Greenshot 自身的设置中调整；RelaxCat 启动集成时读取保存目录。修改保存目录后在 RelaxCat 重新应用集成。
4. 文件夹模式不启动 Greenshot，可配合其他截图软件。只监听指定目录本层的 PNG/JPEG/BMP；启动时已有文件不会批量导入。新建、改写和原子重命名的图片等待稳定且可读取后进入现有队列，默认勾选导出；重复通知不会重复导入，暂时失败会重试。手动导入/Excel 导出忙碌时排队，完成后继续导入。
5. 修改来源、路径、快捷键后重新点击「应用并启动」。点击「停止监控」或选择「关闭」并应用，会解除监听并停止自有 Greenshot。启动过程中也可点击停止取消启动。关闭 RelaxCat 会先取消当前工作、停止监控、清理自有进程树；截图文件会保留。

Greenshot 每个用户会话只允许运行一个实例。如果其他 Greenshot 已运行，RelaxCat 会明确报错，保留其他实例；请手动退出它，或改用仅监控文件夹。Greenshot 意外退出会显示错误并停止监控，可重新应用启动。配置会自动保存。已选择集成时，下次启动 RelaxCat 会自动启动 Greenshot 并监控。

图片原始像素不重采样，DPI 使用统一倍率参与尺寸计算，宽高保持源图片比例。重新规范 Excel 行列尺寸时也会保护已有图片的尺寸，并保留随单元格移动和缩放的设置。默认一图一行；节省空间模式按设计宽度换行。

在「设置 → Excel 排版」中：

- 「起始行」「起始列」指定第一张图片所在的单元格，默认第 1 行、第 1 列（A1），图片对齐单元格左上角。行列从 1 开始计数。每次导出先清空旧图片，再从配置的起始行开始；起始列也用于后续每行的左边界。设置自动保存，下次导出会按新设置替换旧图片。旧版 XLS 工作表按其实际行列上限检查。

- 水平和垂直间距可以分别选择「磅（pt）」或「单元格」，支持小数。1 个水平单元格对应规范化后的实际列宽，1 个纵向单元格对应实际行高。例如行高 18 pt 时，纵向 1.5 个单元格为 27 pt。列宽仍使用 Excel 字符单位，不能直接把 2.38 字符当成 2.38 pt；转换由 Excel 实际测量。
- 「显示工作表网格线」控制目标工作表的网格线显示，默认开启，随工作簿保存。
- 勾选「导出后选中工作表全部图片」，成功导出后会在 Excel 中选中目标工作表的所有图片，不选中非图片对象，也不修改首页图片队列的勾选状态。默认关闭；取消或失败不执行全选。

规则按「单行条件 + 下一行 scale 赋值」成对填写，例如：

```text
width >= 1920 or height >= 1080
scale = 0.5

width < 1000 and height < 800
scale = 0.9
```

支持 `width`、`height`（原始像素）、`dpi`、`dpiX`、`dpiY`、`aspectRatio`；比较运算符 `< <= > >= == !=`；`and` 优先于 `or`，暂不支持括号。首条命中即生效，未命中时保持比例适配设计目标，并限制回退放大倍数。

「取消」在当前图片/COM 操作完成后生效，保留已完成的部分。再次导出会先清空目标工作表中的图片，再重新导出当前勾选的图片。源文件需保留至导出或复制完成。

## Excel 样式优化

离线分析 `.xlsx` / `.xlsm` 的样式表、引用关系和预计输出，不依赖 Excel。`.xls` 和加密工作簿不支持。源路径不能作为输出路径；请始终保留原件，使用独立输出目录。

### 选择处理模式

| 模式 | 开关 | 行为与边界 |
| --- | --- | --- |
| 默认优化 | 无 | 在已覆盖的引用范围内合并等价定义；保留命名样式身份及其基础 XF 所属关系。未知依赖会阻止索引改写。 |
| 未引用样式回收 | `--gc` | 默认关闭。仅回收已知引用图内不可达的定义；命名样式仍作为保留根。 |
| 复制名称清理 | `--remove-copied-named-styles` | 默认关闭。删除满足严格候选条件的复制名称；外部模板或程序仍可能依赖这些名称，必须先审核候选。 |
| 破坏性样式清理 | `--destructive-styles` | 默认关闭的有损实验功能。除 Normal 外删除命名样式与基础 XF 身份，断开继承，并执行回收。可能改变当前外观、保护、宏与后续编辑行为。 |

**有损输出的 `Verified=true` 仅表示选定结构校验通过，不表示外观或行为等价。** Microsoft Excel 原生打开、保存、重开及打印的完整验收仍待完成。模式细节见 [样式优化指南](docs/excel-style-optimizer.md) 和 [有损模式说明](docs/destructive-style-cleanup.md)。新增兼容处理及其适用版本见 [Office 元数据兼容验证](docs/destructive-compatibility-validation.md)。

### 单文件与批量使用

1. 打开「Excel 优化」，选择工作簿，确认模式后点击「分析」。查看 `styles.xml` 未压缩大小、各分区数量、预计删除数和跳过原因；“重复”或“未引用”数量不等于可删除数量。
2. 查看工作表样式预览，再选择「优化并另存为…」。原文件不改写；单文件输出位置若已有文件，可能被替换，请选新路径。
3. 完成后检查实际变化、校验范围及警告，并在目标 Excel 版本上打开输出进行复核。无可执行变化不代表失败，也不能把 ZIP 重新压缩算作样式删除。
4. 单文件确认符合预期后再使用「批量优化」。默认递归，保留相对子目录，输出名追加 `-optimized`；已存在输出默认跳过，显式勾选覆盖后才替换。逐文件显示结果，单个失败后继续，其余已完成结果保留。

分析与样式预览不展示单元格值、公式或批注。**完整报告、日志和截图仍可能包含路径、工作表名、命名样式名、格式代码或扩展标识**，分享前需审核。空白单元格不代表工作簿已完全去除敏感信息。

### 命令行

```powershell
# 只读分析；--json 输出机器可读结果
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- analyze path\to\book.xlsx --json

# 默认模式，输出目录须已存在，输出文件请选择新路径
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize path\to\book.xlsx --output path\to\book.optimized.xlsx

# 批量处理，默认递归且跳过已有输出
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- batch path\to\input --output-dir path\to\output
```

三个命令均支持上表中的模式开关。保留 `.xlsx` / `.xlsm` 输入扩展名；宏二进制保留不代表宏行为已验证。遇到拒绝访问、跳过或无明显缩减时，按 [故障排查与验收](docs/excel-style-optimizer.md#故障排查) 检查原因，不要关闭系统保护来强行写入。

## 验证

```powershell
dotnet build RelaxCat.sln -c Release
dotnet test RelaxCat.sln -c Release --no-build
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build
# 只跑 Excel 样式优化页面流程：生成 artifacts/excel-tool-sample.xlsx、执行分析/优化并截图
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel-tool-only
# 只跑 Excel 批量优化页面流程
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel-batch-only
# 可选：实际启动测试专用 Excel，导出并验证，然后关闭该测试实例
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel
# 可选：真实 Greenshot 截图链路、切换、异常退出、关闭清理
# 需要没有其他 Greenshot 正在运行；测试恢复鼠标原位置，只清理自己的进程
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --greenshot --excel
```

以上是验证命令，不是本分支全部通过的声明。Core 测试包含 Excel 样式分析、去重映射、包保留、合并单元格与批量扫描/继续处理的用例；无界面冒烟测试用于检查页面与模型流程，不能代替 Windows 桌面与 Excel 实测。版本化验证记录见 [样式引用图验证](docs/excel-style-graph-validation.md) 和 [有损模式验证](docs/destructive-style-cleanup.md#verification-performed)。直接集成 Windows 原生截图工具及更多图片格式仍待后续实现。产品意图入口为 `.singularity/README.md`，截图集成验证记录为 `docs/capture-validation.md`，Excel 样式优化记录为 `docs/excel-style-optimizer.md`。

## 许可证

本项目采用 [The Unlicense](LICENSE)，将作品贡献到公共领域：可自由复制、修改、分发、商用或闭源再分发，无需署名。软件按现状提供，不附带任何担保。

随应用分发的 Greenshot 是独立的第三方程序，遵循 **GPL v3**，不受上述授权影响。分发时须保留其许可证与对应的未修改源码，详见 [`vendor/greenshot/README.md`](vendor/greenshot/README.md)。界面与图像处理依赖（Avalonia、SkiaSharp 等）以 MIT 许可发布，具体版本与许可声明见各自的 NuGet 包。

界面翻译维护在 `src/RelaxCat.App/Localization/` 下的 `zh-CN.json`、`ja-JP.json`、`en-US.json`、`ko-KR.json`，构建/发布时复制到应用目录。语言包格式、回退行为及校验命令见 [本地化说明](docs/localization.md)。
