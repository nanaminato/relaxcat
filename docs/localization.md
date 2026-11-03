# 外部 JSON 本地化

参考 RelaxKonOS 的稳定键、外部语言包、语言切换通知及回退方式，RelaxCat 将文案维护在：

- `src/RelaxCat.App/Localization/zh-CN.json`：简体中文基线。
- `src/RelaxCat.App/Localization/ja-JP.json`：日文。
- `src/RelaxCat.App/Localization/en-US.json`：English（英语）。
- `src/RelaxCat.App/Localization/ko-KR.json`：한국어（韩语）。

四份语言包均包含完整的 281 个翻译键，覆盖图片工作区、设置、截图监控、Excel 单文件/批量优化和诊断结果。设置页按中文、日文、英语、韩语顺序显示。

应用启动时扫描 `Localization/*.json`，语言列表完全由文件内容决定。每个文件使用 UTF-8，结构如下（例如 `ko-KR.json`）：

```json
{
  "culture": "ko-KR",
  "displayName": "한국어",
  "sortOrder": 30,
  "strings": {
    "Workspace": "이미지 작업 공간",
    "Ready": "준비됨"
  }
}
```

`culture` 必须是有效语言代码并与文件名对应；`displayName` 是设置页显示的语言自称；可选 `sortOrder` 默认为 100，按数值、语言代码排序。完整语言包可复制现有 `zh-CN.json`，修改元数据并翻译 `strings` 内的值。添加英语则创建 `en-US.json`，指定 `displayName: "English"`。将文件放到应用目录的 `Localization` 下并重启，设置页会自动出现该语言，无需修改 C# 或 XAML。源码目录中新加的 JSON 也会自动进入构建/发布输出。

运行时允许部分翻译，缺失项回退到中文；正式随项目分发的语言包应补全所有键，以通过完整性校验。保持已有稳定键，新增文案时更新所有语言文件。占位符编号（例如 `{0}`、`{1}`）必须对应，顺序可以不同。工作表名称、文件路径等用户数据不翻译。

`UiText` 继续提供现有索引器供 XAML/C# 使用，切换语言会通知绑定并刷新派生显示值。语言包从 `AppContext.BaseDirectory/Localization` 加载并缓存；运行时修改后重启应用生效。构建、发布及单文件发布都会将 JSON 作为外部文件复制，不把翻译嵌入 C#。任意语言缺键、空白或格式参数无效时回退到中文；基线也缺键时显示稳定键。元数据/JSON 损坏的非基线语言包会跳过；已保存的语言包被移除时回退到中文。中文基线文件属于必需的部署文件，缺失/损坏会明确报错。

验证：

```powershell
python tools/verify-localization.py
dotnet run --project tests/RelaxCat.App.SmokeTests -- --excel-tool-only
```

脚本检查重复键、非空字符串、所有已发现语言包的元数据、键集与占位符编号一致性，以及 C#/XAML 的字面量资源引用。动态拼接键由现有功能冒烟检查覆盖。桌面检查读取实际构建输出的 JSON，验证全部资源、即时切换和回退行为。

Excel 页的工作表标签栏和样式网格使用独立滚动条空间，关闭自动隐藏覆盖模式；无溢出时不显示滚动条。布局回归覆盖 40 个长名称工作表、400 个样式、最后标签可达性，以及明暗主题截图。

语言偏好以稳定 `LanguageCulture` 代码写入设置（schema 2），不依赖列表索引。旧设置中的 0/1 仍分别迁移为 zh-CN/ja-JP；新增或调整语言包排序不会改变已保存的语言。资源加载器只缓存当前进程的包，修改后需要重启。

回归检查使用临时英语/韩语包，验证仅添加 JSON 后的发现、真实设置页下拉框切换、语言代码保存、排序变化、旧设置迁移及移除语言包后的回退。完整英语/韩语语言包同时参与构建输出资源检查与 Excel 页面切换、指标和样式详情刷新检查。
