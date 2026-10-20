# Greenshot 与文件夹监控验证

日期：2026-10-20。沿用已有 Confirmed S-CAPTURE-001/002/003、S-IMPORT-001/002、S-UI-002、S-I18N-001；本轮未新增或修改确认意图。实现决策记录为 D-PHASE-002、D-GREENSHOT-001、D-GREENSHOT-LIFECYCLE-001、D-CAPTURE-MONITOR-001、D-CAPTURE-KEYBOARD-001。

## 已验证

- Release 构建：0 警告、0 错误；核心测试 43 项通过，包括快捷键解析、文件写入锁等待、稳定等待、重复通知、临时文件原子重命名、新版本重导入、失败重试和停止时取消。
- 官方 GitHub 最新合适 Windows portable pre-release 为 `v1.4.371-g78211b16d9`，本轮实际下载的原始 ZIP SHA-256 与 API digest 一致：`b915f7bc8dfde4cdfa271f2731bdc1db9c262a3723e6898efe5c8b3e6385103c`。包适用于 x64 Windows，依赖 .NET Framework 4.8。原包、上游许可证和精确源码链接保留于 `vendor/greenshot/`。
- 从 XAML「截图与监控」页启动/停止真实文件夹监控；使用中文与空格目录，确认已有文件不导入、新文件自动进入现有图片队列、忙碌期间等待而不丢图、同版本不重复、停止后不继续导入。
- `--greenshot` 实际启动随包 Greenshot，生成并读取独立 INI，通过 F24 全屏截图、自动落盘、监控解码、队列加入形成真实纵向链路。可选鼠标 hook 把鼠标移动到 (30,40)，同一个 F24 仍被 Greenshot 接收并完成截图。测试结束恢复鼠标原位置。
- 真实检查第二个 provider 遇到已有 Greenshot 时拒绝接管；来源切换停止自有进程；启动取消不残留；模拟子进程异常退出会停止监控并显示错误，可重新启动；关闭主窗口会停止 Greenshot。导入中关闭会等待当前操作取消并完成退出，无需再点一次关闭。
- 发现并修复：上游 CLI 在服务未就绪时会自启动另一个 Greenshot，原就绪探测因此留下额外实例。现改为无命令管道连接并验证 server PID，另用 Job Object 确保自有进程树清理；不使用全局退出命令。
- XAML 保持中日双语与明暗主题，截图详情仍位于独立设置页，操作按钮固定在滚动区外。检查 `artifacts/settings-capture.png`。
- Excel 的真实 COM 插入、比例保护、单元格间距、全选、工作簿复用和模板恢复回归通过。
- 发布目录复制便携包与附属说明，并重新检查发布包 ZIP 摘要。

## 运行

```powershell
dotnet build RelaxCat.sln -c Release
dotnet test RelaxCat.sln -c Release --no-build
dotnet run --project tests/RelaxCat.App.SmokeTests -c Release --no-build -- --greenshot --excel
dotnet publish src/RelaxCat.App -c Release -o artifacts/publish
```

真实 Greenshot 检查需要当前没有其他 Greenshot，测试不会退出用户已有的实例。测试使用专用 F22/F23/F24，不修改用户的 PrintScreen 设置，输出仅保存在本地 artifacts 中。

## 意图核对与边界

- S-CAPTURE-001：选择来源、启动、切换、停止、意外退出和宿主退出均有实际生命周期控制；只管理自己的进程。
- S-CAPTURE-002：同步移动鼠标后继续调用下一 hook，没有吞键、异步延后或额外合成用户按键；测试键注入仅在测试工具中。
- S-CAPTURE-003：原始官方 Windows 便携预发行包随构建与发布输出分发，不依赖机器级 Greenshot 安装；此版本明确为 x64/.NET 4.8。
- S-IMPORT-001/002：监控与手工导入共用工作流，新图片默认参与导出，原有逐项操作与导出筛选继续有效。

设置仍仅在会话内有效，默认不启动监控；停止期间新增的历史图片需手工导入。仅处理目录本层的 PNG/JPEG/BMP，不监听子目录。真实测试覆盖全屏截图；交互式区域拖选、多屏负坐标、系统占用快捷键、网络文件夹与长时间高频截图还需现场验证。Windows 原生截图直接集成不在本轮范围。没有需要人裁决的确认意图冲突。
