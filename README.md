# RelaxCat

Windows 图片整理与 Excel 排版工具。提供手动导入、图片卡片与放大预览、DPI 感知缩放、有序规则、逐行/节省空间排列，以及真实 Excel COM 导出。界面使用 Avalonia XAML + C#，支持明亮/黑暗主题和简体中文、日本語切换。

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

图片原始像素不重采样；Excel 使用 points，设计宽高使用 96 DPI 逻辑像素，列宽使用 Excel 字符单位。默认一图一行；节省空间模式按设计宽度换行。

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
```

截图提供器、Greenshot 打包/管理、全局按键、目录监控及配置持久化尚未纳入本阶段。它们的 Confirmed Singularities 未被修改或取消。产品意图入口为 `.singularity/README.md`。
