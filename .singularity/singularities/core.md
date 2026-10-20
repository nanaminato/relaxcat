# RelaxCat Confirmed Singularities

These records capture product intent explicitly stated by the human during the RelaxCat design conversation. They constrain downstream implementation but intentionally do not prescribe ordinary code structure.

## S-CAPTURE-001

Status: Confirmed

Statement: RelaxCat supports selectable screenshot providers. Greenshot is a bundled executable provider; when selected, RelaxCat manages the Greenshot process when starting, switching providers, and exiting.

Confirmation basis: Human clarification, 2026-10-20.

## S-CAPTURE-002

Status: Confirmed

Statement: For configured monitored keys, RelaxCat may move the mouse to a configured position before the same key continues to Greenshot; RelaxCat's monitoring must not prevent Greenshot from processing that key.

Confirmation basis: Human clarification, 2026-10-20.

## S-CAPTURE-003

Status: Confirmed

Statement: The bundled Greenshot provider targets Windows 10 and later and uses the latest suitable Windows portable pre-release build published by the upstream Greenshot project on GitHub, rather than requiring a machine-wide Greenshot installation. The exact upstream build is a packaging dependency and may be updated independently without changing this intent.

Confirmation basis: Human instruction, 2026-10-20.

## S-IMPORT-001

Status: Confirmed

Statement: Screenshots produced through the monitored capture folder and images imported manually enter the RelaxCat image/export workflow. Long import and export operations expose progress and can be cancelled.

Confirmation basis: Human project description, 2026-10-20.

## S-SIZE-001

Status: Confirmed

Statement: RelaxCat does not normally resample image pixels for Excel layout. It changes the displayed image size, and DPI participates in sizing so images with different DPI metadata have more consistent displayed results.

Confirmation basis: Human clarification, 2026-10-20.

## S-SIZE-002

Status: Confirmed

Statement: Image scaling first evaluates user-defined ordered rules. Conditions are single-line expressions that may combine comparisons with `and` and `or`; rules are tried in order and the first match supplies the scale. If no rule matches, sizing falls back to the configured design size and design scale.

Confirmation basis: Human clarification, 2026-10-20.

## S-SIZE-003

Status: Confirmed

Statement: Fallback sizing preserves aspect ratio, approaches the configured design-size/design-scale target without exceeding its width or height, and limits enlargement so small images are not excessively enlarged.

Confirmation basis: Human clarification, 2026-10-20.

## S-CACHE-001

Status: Confirmed

Statement: Scale calculation should be cacheable because a given effective image size and sizing configuration normally produce a deterministic scale; the cache exists to avoid repeated calculation.

Confirmation basis: Human clarification, 2026-10-20.

## S-LAYOUT-001

Status: Confirmed

Statement: Default Excel layout places one image per row. In space-saving mode, a row still has a maximum usable width derived from design size and design scale; smaller images continue horizontally with configured spacing until the next image no longer fits, then layout continues on the next row.

Confirmation basis: Human clarification, 2026-10-20.

## S-EXCEL-001

Status: Confirmed

Statement: Excel export reuses an available document/session until it is no longer usable. If no usable document exists, RelaxCat creates a new document from the configured template.

Confirmation basis: Human project description and clarification, 2026-10-20.

## S-EXCEL-002

Status: Confirmed

Statement: Before exporting images, RelaxCat normalizes the target worksheet according to configuration, including row heights and column widths. This normalization occurs for both newly created and already available documents.

Confirmation basis: Human clarification, 2026-10-20.

## S-EXCEL-003

Status: Confirmed

Statement: Exported Excel images are configured to move and resize with cells (`xlMoveAndSize`).

Confirmation basis: Human project description, 2026-10-20.

## S-EXCEL-004

Status: Confirmed

Statement: Excel automation uses the Microsoft Office/Excel COM interop model and must remain usable with the Office generations already validated by the human: Excel 2007, Excel 2016, and Excel 2024. The implementation must not intentionally bind itself to behavior that only works with Excel 2016.

Confirmation basis: Human compatibility observation and instruction, 2026-10-20.

## S-UI-001

Status: Confirmed

Statement: RelaxCat's desktop user interface follows an IntelliJ IDEA-style visual and interaction direction. This is a product-level design direction rather than a requirement to reproduce JetBrains UI pixel-for-pixel; implementation details, controls, theming, and platform adaptations may be derived as needed while preserving the recognizable IDEA-style desktop-tool character.

Confirmation basis: Human instruction, 2026-10-20.

## S-I18N-001

Status: Confirmed

Statement: RelaxCat provides both Simplified Chinese and Japanese user-interface localization. User-facing text must be designed so these two languages can be maintained as supported product languages rather than being hard-coded to a single language.

Confirmation basis: Human instruction, 2026-10-20.

## S-UI-002

Status: Confirmed

Statement: Implement desktop views with XAML plus C#. The home page exposes only a small set of common settings; more detailed settings belong on separate settings pages.

Confirmation basis: Human instruction, 2026-10-20: “使用xaml+cs，而不是纯cs，另外首页不要放置太多设置，放置部分即可。同时更多的细节使用单独的页来设置”。

## S-UI-003

Status: Confirmed

Statement: Provide both light and dark interface themes.

Confirmation basis: Human instruction, 2026-10-20: “同时添加明亮和黑暗主题”。

## S-IMPORT-002

Status: Confirmed

Statement: The home page displays each imported image individually. Each image can be removed and copied, and has a checkbox controlling whether it is included in export. Provide an elegant image preview experience.

Confirmation basis: Human instruction, 2026-10-20: “主页时显示的图片，每一个都能显示出来并且对单个的可以进行移除和复制，同时包含一个checkbox，他会决定是否会出现到导出中，预览时提供优雅的预览效果”。

## S-EXCEL-005

Status: Confirmed

Statement: Provide an option to select all pictures in Excel after export.

Confirmation basis: Human instruction, 2026-10-20: “添加选项-导出后选择所有图片”，alongside the Excel settings screenshot.

## S-LAYOUT-002

Status: Confirmed

Statement: Horizontal and vertical image gaps also support cell units.

Confirmation basis: Human instruction, 2026-10-20: “横向和纵向间隔添加 单元格单位”。

## S-SIZE-004

Status: Confirmed

Statement: Export preserves the source image's aspect ratio; RelaxCat must not stretch pictures during sizing, insertion or worksheet normalization.

Confirmation basis: Human correction, 2026-10-20: “为什么现在导出的图片的比例会发生变化，这个不应该”。This reinforces S-SIZE-003 without removing S-EXCEL-003.

## S-SETTINGS-001

Status: Confirmed

Statement: RelaxCat settings persist across application restarts, including cell gap units and capture provider selection.

Confirmation basis: Human correction, 2026-10-20: settings appear not to persist.

## S-CAPTURE-004

Status: Confirmed

Statement: Greenshot owns its settings; RelaxCat may read selected values but must preserve them on subsequent launches. When integration is selected, application startup launches Greenshot and starts monitoring automatically.

Confirmation basis: Human instruction, 2026-10-20.

## S-EXCEL-006

Status: Confirmed

Statement: Provide a persistent setting controlling whether the exported worksheet displays gridlines.

Confirmation basis: Human instruction, 2026-10-20.

## S-IMPORT-003

Status: Confirmed

Statement: Provide a clear-all action that removes all images from the workspace. Source files and previously exported Excel content are retained.

Confirmation basis: Human instruction, 2026-10-20: add an all-remove button to remove images from the workspace.

## S-IMPORT-004

Status: Confirmed

Statement: Support automatic chronological sorting, including automatically ordering images after import.

Confirmation basis: Human instruction, 2026-10-20.

## S-WORKFLOW-001

Status: Confirmed

Statement: Preserve the current workflow for this improvement pass. Do not add project saving, export-only-unexported behavior or automatic export.

Confirmation basis: Human clarification, 2026-10-20.
