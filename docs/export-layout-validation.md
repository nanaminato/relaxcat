# Excel 间距、全选与比例修复验证

日期：2026-10-20。新增人类明确确认的 S-EXCEL-005、S-LAYOUT-002、S-SIZE-004；保留既有 xlMoveAndSize、行列规范化及 Office 兼容性意图。推导决策见 D-UNITS-002、D-ASPECT-001、D-EXCEL-SELECTION-001。

## 根因与修复

1. 原计算分别用 X/Y DPI 除宽高，导致 DPI 不相同时改变像素比例。回归先复现了 1.5 的比例变为 3、0.75、6.25；修复使用统一有效 DPI，三个案例均恢复 1.5。
2. Excel 会按行高、列宽分别缩放 xlMoveAndSize 图片。真实 COM 回归把列宽设为截图中的 2.38 字符时，旧实现将约 609.88×405.00 pt 改成 168.13×560.25 pt。现在规范化期间临时冻结图片并恢复几何尺寸，之后仍为 609.88×405.00 pt，Placement=1 和 LockAspectRatio=-1 均通过检查。

已被旧版本拉伸的图片没有可靠的源图映射，修复不声称能自动复原它们；请移除旧图片后重新导出。用户在 Excel 中单独调整行列尺寸，仍受已确认的 xlMoveAndSize 行为影响。

## 验证范围

- Release 构建：0 警告，0 错误。
- 核心测试：30 项通过，新增不等轴 DPI、独立间距单位、小数单元格、跨批次间距、全选开关、失败/取消/空批次不全选、非法单位拒绝等验证。
- XAML 交互检查：横纵单位和全选开关能进入真实 ExportOptions，切换语言不丢失所选单位；保留首页卡片、主题、预览、复制与勾选导出的回归。检查 `artifacts/settings-excel.png` 的新设置布局。
- 真实 Excel（报告 16.0）：测得 2.38 字符列宽对应本机 18 pt，而不是 2.38 pt；分别验证 1.5 列的水平间距、2 行的垂直间距、追加间距。启用全选后检查 ShapeRange 包括此前全部图片，排除测试矩形；插入、工作簿复用、模板恢复和单元格随动设置也通过。
- 兼容性仍采用 Excel 2007 已有 API；各代 Office 的独立安装环境尚未逐个验证。

运行：

```powershell
dotnet build RelaxCat.sln -c Release
dotnet test RelaxCat.sln -c Release --no-build
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --excel
```
