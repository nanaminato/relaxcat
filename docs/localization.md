# 外部 JSON 本地化

参考 RelaxKonOS 的稳定键、外部语言包、语言切换通知及回退方式，RelaxCat 将文案维护在：

- `src/RelaxCat.App/Localization/zh-CN.json`：简体中文基线。
- `src/RelaxCat.App/Localization/ja-JP.json`：日文。

两份文件均为 UTF-8 的扁平 JSON 对象：`"Workspace": "图片工作台"`。保持已有稳定键，新增文案时同时更新两个语言文件。占位符编号（例如 `{0}`、`{1}`）必须对应，顺序可以不同。工作表名称、文件路径等用户数据不翻译。

`UiText` 继续提供现有索引器供 XAML/C# 使用，切换语言会通知绑定并刷新派生显示值。语言包从 `AppContext.BaseDirectory/Localization` 加载并缓存；运行时修改后重启应用生效。构建、发布及单文件发布都会将 JSON 作为外部文件复制，不把翻译嵌入 C#。日文缺键、空白、文件缺失/损坏或格式参数无效时回退到中文；基线也缺键时显示稳定键。中文基线文件属于必需的部署文件，缺失/损坏会明确报错。

验证：

```powershell
python tools/verify-localization.py
dotnet run --project tests/RelaxCat.App.SmokeTests -- --excel-tool-only
```

脚本检查重复键、非空字符串、两语键集与占位符编号一致性，以及 C#/XAML 的字面量资源引用。动态拼接键由现有功能冒烟检查覆盖。桌面检查读取实际构建输出的 JSON，验证全部资源、即时切换和回退行为。

Excel 页的工作表标签栏和样式网格使用独立滚动条空间，关闭自动隐藏覆盖模式；无溢出时不显示滚动条。布局回归覆盖 40 个长名称工作表、400 个样式、最后标签可达性，以及明暗主题截图。
