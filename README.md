# RelaxCat

Windows 图片整理与 Excel 排版工具。第一阶段提供手动导入、图片预览、DPI 感知缩放、有序规则、逐行/节省空间排列，以及真实 Excel COM 导出。支持简体中文、日本語切换。

## 启动

需要 Windows 10 或更新版本、.NET 8 Desktop/Runtime，以及构建用 .NET SDK 10.0.300（见 `global.json`）。Excel 导出需要安装桌面版 Microsoft Excel；没有 Excel 时仍可导入与预览。

```powershell
dotnet run --project src/RelaxCat.App -c Release
```

也可启动构建生成的 `src/RelaxCat.App/bin/Release/net8.0-windows/RelaxCat.App.exe`。

1. 点击「导入图片」，选择本地 PNG、JPEG 或 BMP；左侧选择图片，中间显示预览。
2. 右侧配置设计尺寸、缩放、回退最大放大倍数、规则、间距、Excel 行列尺寸及模板路径。右侧可滚动。留空模板使用 Excel 默认模板。
3. 点击「导出到 Excel」。首次创建工作簿，后续导出复用该工作簿并追加图片；如果已关闭则重新创建。导出前会重新规范工作表行高、列宽，可能影响该工作表的原有内容。
4. 在 Excel 中自行保存。关闭 RelaxCat 不会强行关闭已导出的工作簿。

图片原始像素不重采样；Excel 使用 points，设计宽高使用 96 DPI 逻辑像素，列宽使用 Excel 字符单位。默认一图一行；节省空间模式按设计宽度换行。

规则按「单行条件 + 下一行 scale 赋值」成对填写，例如：

```text
width >= 1920 or height >= 1080
scale = 0.5

width < 1000 and height < 800
scale = 0.9
```

支持 `width`、`height`（原始像素）、`dpi`、`dpiX`、`dpiY`、`aspectRatio`；比较运算符 `< <= > >= == !=`；`and` 优先于 `or`，暂不支持括号。首条命中即生效，未命中时保持比例适配设计目标，并限制回退放大倍数。

「取消」在当前图片/COM 操作完成后生效，保留已完成的部分。再次导出会追加整个当前队列，若只需导出余下图片，请先移除已导出的条目。源文件需保留至导出完成。

## 验证

```powershell
dotnet build RelaxCat.sln -c Release
dotnet test RelaxCat.sln -c Release --no-build
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build
# 可选：实际启动测试专用 Excel，导出并验证，然后关闭该测试实例
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel
```

截图提供器、Greenshot 打包/管理、全局按键、目录监控及配置持久化尚未纳入本阶段。它们的 Confirmed Singularities 未被修改或取消。产品意图入口为 `.singularity/README.md`。
