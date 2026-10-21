# RelaxCat 项目长期约定

## Avalonia 布局/样式（踩过的坑，别重复踩）

- 需要“撑满但有最大宽度”时用 `HorizontalAlignment="Stretch"` + `MaxWidth`；`Center` 会按内容宽度收缩（外观页据此把 ComboBox 压成 72px）。
- `ScrollViewer` 内的居中栅格必须同时设 `HorizontalContentAlignment="Stretch"`，否则星号列拿不到宽度。
- `Path` + `Stretch="Uniform"` 会连描边一起缩放（16px 显示 ≈ thickness×1.3）；几何形状偏“小块堆叠”时最容易糊成一团。
- TabControl 的页面卡片只包内容区：样式打在 `/template/ ContentPresenter#PART_SelectedContentHost`（Background/Border/CornerRadius/VerticalAlignment=Top）。给 TabControl 外套 Border 会裁切选中 Tab 背景。
- 超高页面用 `TabControl.settings ScrollViewer { MaxHeight = $parent[TabControl].Bounds.Height }` 兜住，靠卡片内滚动而不是页面溢出。

## UI 约定

- 图标统一为 `App.axaml` 里的 `<StreamGeometry x:Key="IconXxx">`（16x16 描边路径）+ `Path Classes="glyph"`；描边色按 `Button.xxx` 选择器逐级覆盖，不引外部图标资源，也不要用 template binding。
- 按钮里放图标+文字时，内容是 `StackPanel(Horizontal)`；冒烟测试用 `Label(button)` 读文本，不要断言 `Content` 是字符串。
- 主/次/危险/quiet 四类按钮的全部状态（hover/pressed/disabled）都要在 App.axaml 里显式覆盖控件与 `/template/ ContentPresenter#PART_ContentPresenter`。
- 文案必须中日双语放在 `UiText.cs`；新增界面文本先加键再用 `{Binding Text[...]}`。
- 设置页正文为两栏（`Grid ColumnDefinitions="*,*"`, `ColumnSpacing=40`, `MaxWidth=1040`），窄窗口靠卡片内滚动承载溢出。

## 构建与版本

- 目标框架统一 .NET 10：`RelaxCat.App` / `SmokeTests` 为 `net10.0-windows`，`RelaxCat.Core` / `Core.Tests` 为 `net10.0`；SDK 版本由 `global.json` 钉在 10.0.300。
- 应用版本号的唯一来源是 `RelaxCat.App.csproj` 的 `<Version>`（当前 0.3.0）。不要再用发布命令行 `-p:Version=` 覆盖，否则普通构建产物与 Release 不一致（曾经构建出 1.0.0）。
- 测试机已缓存 `microsoft.netcore.app.runtime.win-x64 10.0.8` 与 host/ref 10.0.x，可离线 `dotnet publish -r win-x64 --self-contained true`。
