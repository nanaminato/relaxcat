# RelaxCat

Windows 截图收集与 Excel 排版工具。提供捆绑 Greenshot 截图、文件夹自动导入、手动导入、图片卡片与放大预览、DPI 感知缩放、有序规则、逐行/节省空间排列，以及真实 Excel COM 导出。界面使用 Avalonia XAML + C#，支持明亮/黑暗主题和简体中文、日本語切换。

## 启动

需要 Windows 10 或更新版本、.NET 8 Desktop/Runtime，以及构建用 .NET SDK 10.0.300（见 `global.json`）。Excel 导出需要安装桌面版 Microsoft Excel；没有 Excel 时仍可导入与预览。

```powershell
dotnet run --project src/RelaxCat.App -c Release
```

也可启动构建生成的 `src/RelaxCat.App/bin/Release/net8.0-windows/RelaxCat.App.exe`。

1. 点击「导入图片」，选择本地 PNG、JPEG 或 BMP。首页以卡片展示每张图片；取消「参与导出」勾选可保留图片但不导出。卡片上的「复制」复制原始图片到剪贴板，「移除」只移出队列，不删除源文件。
2. 点击缩略图打开放大预览：支持适应窗口、缩放与滚动、上一张/下一张，左右方向键切换，Esc 或点击背景关闭。预览中的勾选框与卡片同步。
3. 首页仅放设计缩放与节省空间排列。点击左侧「设置」或「详细设置」，在「尺寸与规则」「Excel 排版」「外观与语言」独立设置页面调整其他选项。空模板使用 Excel 默认模板。顶部「切换明暗」也可快速切换主题。
4. 点击「导出到 Excel」，仅按队列顺序导出已勾选图片。首次创建工作簿，后续复用并追加；如果已关闭则重新创建。导出前会重新规范工作表行高、列宽，可能影响该工作表的原有内容。
5. 在 Excel 中自行保存。关闭 RelaxCat 不会强行关闭已导出的工作簿。主题、语言、设置与图片队列目前均为会话状态，重启后恢复默认。

## 截图与文件夹监控

1. 点击左侧「截图与监控」，选择「Greenshot（随应用捆绑）」或「仅监控文件夹」，指定截图文件夹，然后点击「应用并启动」。默认关闭，不会在未选择时接管系统截图。
2. Greenshot 模式首次启动会从随应用分发的 ZIP 校验并解压官方便携预发行包，无需另装 Greenshot。当前捆绑版本为 `1.4.371-g78211b16d9`，适用于 Windows 10+ x64，需要 .NET Framework 4.8。来源、校验和及上游许可证位于 `vendor/greenshot/`，发布目录也包含这些文件。
3. 默认按 PrintScreen 选择区域，Alt+PrintScreen 截取活动窗口，Ctrl+PrintScreen 截取全屏。可修改为 F1–F24、PrintScreen，或 Ctrl/Alt/Shift/Win 与字母、数字、功能键的组合。三种截图键不能重复，普通字母/数字需要 Ctrl、Alt 或 Win 修饰键。若系统已占用某个按键，可改用其他组合。
4. 「按键监听与鼠标位置」可启用截图前移动鼠标，配置监听快捷键与桌面 X/Y 坐标（负数可用于左侧/上方显示器）。RelaxCat 在同一按键继续交给 Greenshot 前移动鼠标，始终转发按键。
5. 文件夹模式不启动 Greenshot，可配合其他截图软件。只监听指定目录本层的 PNG/JPEG/BMP；启动时已有文件不会批量导入。新建、改写和原子重命名的图片等待稳定且可读取后进入现有队列，默认勾选导出；重复通知不会重复导入，暂时失败会重试。手动导入/Excel 导出忙碌时排队，完成后继续导入。
6. 修改来源、路径、快捷键后重新点击「应用并启动」。点击「停止监控」或选择「关闭」并应用，会解除监听并停止自有 Greenshot。启动过程中也可点击停止取消启动。关闭 RelaxCat 会先取消当前工作、停止监控、清理自有进程树；截图文件会保留。

Greenshot 每个用户会话只允许运行一个实例。如果其他 Greenshot 已运行，RelaxCat 会明确报错，保留其他实例；请手动退出它，或改用仅监控文件夹。Greenshot 意外退出会显示错误并停止监控，可重新应用启动。配置目前仍为会话状态。

图片原始像素不重采样，DPI 使用统一倍率参与尺寸计算，宽高保持源图片比例。重新规范 Excel 行列尺寸时也会保护已有图片的尺寸，并保留随单元格移动和缩放的设置。默认一图一行；节省空间模式按设计宽度换行。

在「设置 → Excel 排版」中：

- 水平和垂直间距可以分别选择「磅（pt）」或「单元格」，支持小数。1 个水平单元格对应规范化后的实际列宽，1 个纵向单元格对应实际行高。例如行高 18 pt 时，纵向 1.5 个单元格为 27 pt。列宽仍使用 Excel 字符单位，不能直接把 2.38 字符当成 2.38 pt；转换由 Excel 实际测量。
- 勾选「导出后选中工作表全部图片」，成功导出后会在 Excel 中选中目标工作表的所有图片（包含以前已导出的图片），不选中非图片对象，也不修改首页图片队列的勾选状态。默认关闭；取消或失败不执行全选。

规则按「单行条件 + 下一行 scale 赋值」成对填写，例如：

```text
width >= 1920 or height >= 1080
scale = 0.5

width < 1000 and height < 800
scale = 0.9
```

支持 `width`、`height`（原始像素）、`dpi`、`dpiX`、`dpiY`、`aspectRatio`；比较运算符 `< <= > >= == !=`；`and` 优先于 `or`，暂不支持括号。首条命中即生效，未命中时保持比例适配设计目标，并限制回退放大倍数。

「取消」在当前图片/COM 操作完成后生效，保留已完成的部分。再次导出会追加当前勾选的图片，若只需导出余下图片，请先取消已完成条目的勾选或将其移除。源文件需保留至导出或复制完成。

## 验证

```powershell
dotnet build RelaxCat.sln -c Release
dotnet test RelaxCat.sln -c Release --no-build
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build
# 可选：实际启动测试专用 Excel，导出并验证，然后关闭该测试实例
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel
# 可选：真实 Greenshot 截图链路、鼠标与按键透传、切换、异常退出、关闭清理
# 需要没有其他 Greenshot 正在运行；测试恢复鼠标原位置，只清理自己的进程
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --greenshot --excel
```

直接集成 Windows 原生截图工具、配置持久化及更多图片格式仍待后续实现。产品意图入口为 `.singularity/README.md`，截图集成验证记录为 `docs/capture-validation.md`。
