# 第一阶段验证记录

日期：2026-10-20。协议：Singularity 0.1.0（按 `.singularity/config.md` 固定来源读取）。初始实现未新增 Confirmed Singularity；本次界面迭代按用户明确指令记录 S-UI-002、S-UI-003、S-IMPORT-002，保留原有确认意图。

## 实现与意图核对

| Confirmed intent | 本阶段实现 / 验证边界 |
| --- | --- |
| S-IMPORT-001 | 本地 PNG/JPEG/BMP 批量导入、进度、逐项错误、取消，导出进度与取消；目录自动导入延期 |
| S-SIZE-001 | 解码读取像素/DPI；沿用纯核心尺寸计算；Excel AddPicture 嵌入源文件并指定显示尺寸，无布局重采样 |
| S-SIZE-002 / 003 | 有序规则、and/or、首条命中；未命中适配设计目标与最大回退放大倍数。修复旧代码对命中规则的额外截断；完整规则校验先于 Excel 修改 |
| S-CACHE-001 | 尺寸/规则/配置作为键，界面复用 exporter/cache；测试区分 DPI、配置和规则 |
| S-LAYOUT-001 | 默认逐行；节省空间按设计宽度左到右排列与换行；points 转换测试 |
| S-EXCEL-001 | 复用 RelaxCat 创建的工作簿；关闭后从配置模板恢复；空模板使用 Excel 默认模板，不接管用户其他 Excel 文档 |
| S-EXCEL-002 | 新建与复用均设置工作表行高/列宽，再读取实际 shape 边界决定追加位置 |
| S-EXCEL-003 | 每张图片设置 Placement=1（xlMoveAndSize）；真实 COM 检查 |
| S-EXCEL-004 | 使用早期 Office 已有的 COM 接口，不引用特定版本 PIA；本机真实 Excel 报告版本 16.0。Excel 2007/2016/2024 各版本独立运行仍需补充，16.0 不能区分 2016 与 2024 |
| S-UI-001 / 002 | Avalonia XAML + C#，IDEA 风格桌面框架、图片卡片工作台、单独设置视图及尺寸/Excel/外观三个设置页；首页只保留设计缩放和紧凑排列 |
| S-UI-003 | 明亮/黑暗主题字典覆盖主界面、设置页与预览；快速切换按钮和设置选项均验证 |
| S-IMPORT-002 | 每张图片独立卡片，带导出 checkbox、复制原图至剪贴板和移除；导出快照仅包含已勾选项。遮罩预览支持淡入、适应窗口、缩放、滚动、导航与 Esc；预览 checkbox 与卡片同步 |
| S-I18N-001 | 中文/日文标签与状态消息；自动切换测试。底层异常作为原始诊断附加显示 |
| S-CAPTURE-001 / 002 / 003 | 本阶段不提供截图提供器，不伪造 Greenshot 实现，也未更改相关确认意图；后续需完整实现生命周期、按键透传和官方便携预发行包 |

## 检查

- `dotnet build RelaxCat.sln -c Release`：通过，0 警告、0 错误。
- `dotnet test RelaxCat.sln -c Release --no-build`：16 项核心测试通过；覆盖规则优先级、DPI、放大限制、排版、导出顺序/单位、取消、缓存、非法数值与规则。
- `dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build`：真实 XAML 加载及控件初始化、页面导航、明暗主题、中日文切换（含已导入卡片及稳定主题选择）、首页设置绑定到导出配置、批量导入、所有图片卡片渲染、逐项勾选过滤实际 exporter 输入、预览 checkbox 同步、复制到无头测试剪贴板并检查原始像素尺寸、移除且源文件保留、预览缩放/导航/Esc、空选择不导出及损坏文件处理。生成 `artifacts/desktop-gallery-light.png`、`desktop-gallery-dark.png`、`desktop-preview-light.png`、`desktop-preview-dark.png` 及各设置页截图，已逐项视觉检查。
- 加 `-- --excel`：启动测试自有 Excel，真实插入图片，验证 points 尺寸、xlMoveAndSize、复用工作簿、重新规范行列、关闭后从生成的 `.xltx` 模板恢复。测试退出关闭自己的 Excel，不修改用户现有文档。
- `git diff --check`：检查补丁空白错误。
- 原生 Windows 启动：实际运行 Release `RelaxCat.App.exe`，确认创建标题为 `RelaxCat` 的主窗口后结束测试进程。除无头渲染检查外，原生窗口启动也已通过。

发现并修复的问题：原有测试缺少 xUnit using；规则命中被回退放大限制覆盖；后续规则/短路分支语法错误未提前验证；非有限尺寸可进入排版；COM FinalRelease 会破坏共享 RCW，改为按获取次数成对 Release。集成测试阻塞桌面同步上下文的问题也已修复。

本轮界面验证发现并修复：切换语言替换主题选项集合会清空选择，改为稳定的主题选项对象，仅更新其翻译标签；预览的适应尺寸误包含容器内边距，改为使用实际图片视口尺寸，避免适应窗口时仍出现滚动条。预览解码结果带请求序号，过期结果及时释放，关闭后不会重新打开。

## 限制

主题、语言、设置与队列暂不持久化；取消保留已完成条目，重试追加当前勾选的图片；大文件取消须等待当前解码/COM 调用完成。Excel 忙碌/模态窗口可能导致调用失败，界面显示错误并允许用户处理后重试。剪贴板的原图尺寸已在无头环境验证，外部应用中的 Windows 粘贴效果仍需人工检查。未验证每个支持的 Office 版本、人工真实按键操作、目录截图链路、Greenshot 打包及长时间压力场景。

没有需要人裁决的确认意图冲突。现有设计中的 Excel-2016 专属提案及原规则截断属于下游证据矛盾，已按 S-EXCEL-004、S-SIZE-002 修正并记录于 Derived Decisions。
