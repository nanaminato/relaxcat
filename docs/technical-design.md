# RelaxCat Technical Design

> Status: Initial design / living document

> Phase-one implementation now follows D-PHASE-001 in `.singularity/decisions/architecture.md`.
> See `README.md` for the runnable workflow and `docs/phase-one-validation.md` for verified scope.
> The suggested phases and boundaries below remain design evidence, not confirmed product intent.

> Current desktop UI: `App.axaml` holds themes/styles; `MainWindow.axaml` holds the gallery and lightbox;
> `Views/SettingsView.axaml` holds separate sizing, Excel, and appearance pages. Observable models provide
> shared settings, per-image export selection and localization. See D-UI-002, D-GALLERY-001 and D-PREVIEW-001.

> Capture integration now implements D-PHASE-002: `Core/Capture` contains stable-folder monitoring and shortcut parsing;
> `App/Capture` contains the pinned portable Greenshot dependency, owned process/job lifecycle, and Win32 keyboard/mouse adapter;
> `MainWindow.Capture.cs` serializes monitored imports with the existing UI workflow. See `docs/capture-validation.md`.

## 1. Overview

RelaxCat is a Windows-first screenshot collection and Excel image layout tool. It is intended as a clean, more extensible reimplementation of an earlier internal tool, without reusing the original source code.

The application collects screenshots or manually imported images, calculates a display scale according to configurable rules, and exports the images into Microsoft Excel with deterministic layout and sizing.

Avalonia is used for the desktop UI. Windows-specific integrations such as global keyboard monitoring, mouse movement, Greenshot lifecycle management, and Microsoft Office COM automation are isolated behind platform/service abstractions.

## 2. Primary workflow

### 2.1 Screenshot capture

1. User selects a capture provider.
2. If the Greenshot provider is selected, RelaxCat starts the bundled `Greenshot.exe` and manages its lifecycle.
3. RelaxCat monitors configured global keys.
4. When a configured key is observed, RelaxCat first moves the mouse pointer to the configured position if automatic mouse movement is enabled.
5. The key is not consumed; Greenshot continues to receive and process the same key event.
6. Greenshot writes the captured image into a configured output folder.
7. RelaxCat monitors that folder, waits until the new file is completely written and accessible, then imports it.
8. The imported image enters the same image pipeline as a manually imported image.

Windows 10/native capture is also considered a selectable capture provider. Capture-provider-specific behavior must not leak into the image/export core.

### 2.2 Manual import

Users may manually import one or multiple images. Large imports expose progress and support cancellation. Manually imported and automatically captured images use the same metadata, scale calculation, layout, and Excel export pipeline.

### 2.3 Excel export

1. Resolve the current Excel document/session.
2. If no usable workbook exists, create a new workbook from the configured template.
3. Whether the workbook is newly created or already exists, normalize the target worksheet before export according to the active configuration.
4. Normalization includes configured row heights and column widths.
5. Calculate image display size and placement.
6. Insert images using Office/Excel automation.
7. Configure inserted shapes to move and resize with cells (`xlMoveAndSize`).
8. Report export progress and allow cancellation between safe operation boundaries.
9. Reuse the Excel application/workbook while it remains usable. If the document is closed or becomes invalid, resolve/create a usable document on the next export.

## 3. Architecture

Suggested solution boundaries:

```text
RelaxCat
├─ RelaxCat.App                 Avalonia application / composition root
├─ RelaxCat.UI                  Views, ViewModels, UI services
├─ RelaxCat.Core                Domain models and application logic
├─ RelaxCat.Image               Image metadata, sizing, scale rules and cache
├─ RelaxCat.Excel               Excel layout abstractions and export orchestration
└─ RelaxCat.Windows             Win32, COM, Greenshot and Windows capture integrations
```

The exact number of projects may initially be smaller. The important boundary is that Core/UI must not directly own COM objects, Win32 hooks, or Greenshot processes.

Key service abstractions may include:

```text
ICaptureProvider
IKeyboardMonitor
IMouseController
ICaptureFolderMonitor
IImageMetadataReader
IScaleRuleEngine
IScaleCache
IImageLayoutEngine
IExcelSession
IExcelExporter
IWorkbookNormalizer
```

## 4. Capture providers

### 4.1 Greenshot provider

Greenshot is bundled with RelaxCat as an executable dependency.

Responsibilities:

- Start Greenshot when the provider becomes active.
- Detect/manage the RelaxCat-owned Greenshot process.
- Configure or coordinate its output folder.
- Stop/clean up the managed process when switching away from Greenshot or exiting RelaxCat.
- Monitor the output folder for newly created captures.
- Avoid importing partially written files.
- Avoid duplicate events from `FileSystemWatcher` by de-duplicating file identity/path and waiting for file stability.

RelaxCat must distinguish a process it started from unrelated Greenshot instances where practical, so application shutdown does not terminate arbitrary user processes.

### 4.2 Keyboard and mouse ordering

A configured global key event follows this ordering:

```text
Key observed
  -> evaluate configured action
  -> move pointer to configured (X, Y), if enabled
  -> do not consume the key
  -> Greenshot receives the same key
  -> capture proceeds
```

The implementation must minimize work inside the low-level hook callback. Configuration should be available from an immutable/current snapshot and expensive work should be dispatched safely.

Because ordering is time-sensitive, this area requires integration tests/manual timing tests. The system should avoid artificial delays unless a delay proves necessary and is configurable.

## 5. Image model

An imported image should retain at least:

```text
SourcePath
PixelWidth
PixelHeight
DpiX
DpiY
AspectRatio
ImportSource (Manual / Greenshot / NativeCapture / ...)
ImportedAt
CalculatedScale
CalculatedDisplayWidth
CalculatedDisplayHeight
```

RelaxCat does not resample the source image as part of normal layout. Scaling changes the display dimensions of the Excel Shape.

## 6. DPI normalization

Applying the same nominal scale directly to images with different DPI metadata can produce inconsistent physical/display sizes in Office. RelaxCat therefore includes DPI in its size calculation so equivalent source images produce more consistent Excel output.

The sizing engine must keep these concepts separate:

- pixel dimensions;
- source DPI;
- logical/design dimensions;
- Excel display dimensions (points);
- scale selected by rules or fallback sizing.

DPI normalization affects display sizing; it does not require image resampling.

The current implementation uses one effective DPI (the mean of the validated X/Y metadata) for both dimensions, so unequal DPI axes never distort the source pixel aspect ratio. This corrects the former independent-axis conversion; see S-SIZE-004 and D-ASPECT-001.

## 7. Design size and fallback scaling

Configuration defines a design width and design height plus a design scale.

The design size represents the reference image/design area. The design scale describes the desired displayed proportion of that reference size.

If no scale rule matches an image, RelaxCat calculates a scale that makes the image as close as possible to the configured target design size while:

- preserving the original aspect ratio;
- never exceeding the target width;
- never exceeding the target height;
- respecting the configured maximum enlargement ratio.

The maximum enlargement ratio prevents very small source images from being enlarged enough to become visibly poor.

Conceptually:

```text
targetWidth  = designWidth  * designScale
targetHeight = designHeight * designScale

fitScale = min(targetWidth / normalizedWidth,
               targetHeight / normalizedHeight)

finalScale = min(fitScale, maximumEnlargement)
```

The production calculation must account for DPI normalization and Excel point conversion rather than mixing pixels and Office units directly.

## 8. Scale rule language

Before fallback design-size calculation, RelaxCat evaluates user-defined scale rules in declaration order.

Example:

```text
width > 1 and width < 1000
scale = 0.9

width >= 1000 and width < 1920
scale = 0.75

width >= 1920 or height >= 1080
scale = 0.5
```

Rules are evaluated sequentially. The first matching rule supplies the scale. If no rule matches, fallback design-size scaling is used.

Initial language requirements:

- one condition expression per rule;
- multiple comparisons in a condition;
- `and` and `or` logical operators;
- comparison operators such as `<`, `<=`, `>`, `>=`, `==`, `!=`;
- numeric values;
- variables such as `width` and `height`;
- `scale` result assignment;
- deterministic first-match behavior.

Likely future variables include `dpi`, `dpiX`, `dpiY`, and `aspectRatio`.

The rule language should be implemented as a small dedicated parser/evaluator rather than executing arbitrary C# or scripts. This keeps configuration deterministic, cacheable, testable, and safe.

A simple AST is sufficient:

```text
Rule
├─ Condition
│  ├─ Comparison
│  ├─ And
│  └─ Or
└─ Scale
```

Operator precedence must be documented. Recommended initial precedence is comparison > `and` > `or`, with parentheses added either initially or as an early extension.

## 9. Scale cache

For the same effective input, scale calculation is deterministic. A cache avoids repeatedly parsing/evaluating rules and recalculating equivalent image dimensions.

The cache key must represent every input capable of changing the result, for example:

```text
PixelWidth
PixelHeight
DpiX
DpiY
ScaleRuleSetVersion/Hash
DesignWidth
DesignHeight
DesignScale
MaximumEnlargement
```

The cache value can contain the selected scale and derived normalized/display dimensions.

Configuration changes that affect sizing must invalidate the corresponding cached results. A hash/version of the effective sizing configuration is preferable to manually clearing unrelated cache entries.

Two cache layers may eventually be useful:

1. parsed-rule cache: rule text -> parsed AST;
2. scale-result cache: image dimensions + effective sizing configuration -> calculated result.

## 10. Excel layout

### 10.1 Default mode

By default, one image occupies one layout row. Images advance vertically using configured vertical spacing.

### 10.2 Space-saving mode

When space-saving mode is enabled, a row may contain multiple images.

Each row has a maximum usable design width derived from the configured design size and design scale. Images are placed from left to right with the configured horizontal interval.

For each image:

```text
if row is empty:
    place image
else if currentWidth + horizontalGap + imageWidth <= maximumRowWidth:
    place image in current row
else:
    begin next row
    place image there
```

The row height is determined by the tallest displayed image in that row. The next row begins after that height plus the configured vertical interval.

This is intentionally a deterministic shelf/row packing algorithm rather than an aggressive bin-packing algorithm. Stable, predictable Excel output is more important than squeezing every possible gap.

## 11. Worksheet normalization

Every export normalizes the destination worksheet, including when the workbook already exists.

Normalization includes at least:

- configured column widths;
- configured row heights;
- establishing the configured starting X/Y or starting cell/placement origin;
- ensuring the sheet is in a known state required by the layout algorithm.

A newly created workbook is created from the configured template and then normalized before the first image is inserted.

Normalization and image insertion should be separate operations so normalization behavior can be tested and changed independently.

Before changing cell dimensions, temporarily freeze existing pictures and snapshot their geometry, then restore the geometry, aspect lock and original placement even on failure. Exported pictures finish with `xlMoveAndSize` and locked aspect ratio. Read the normalized cell's physical Width/Height for optional cell-unit gaps, rather than treating ColumnWidth characters as points. After successful export, an optional Shapes.Range selection selects all worksheet pictures, excluding non-picture objects.

## 12. Excel integration

The first implementation uses late-bound Microsoft Excel/Office COM automation with APIs available in Excel 2007, 2016 and 2024, as required by S-EXCEL-004. Actual version coverage must be reported separately from API compatibility.

COM/Interop remains appropriate because RelaxCat needs live Excel behavior rather than only producing an `.xlsx` package:

- reuse a running Excel session/workbook;
- create workbooks from templates;
- modify row/column dimensions;
- insert and position Shapes;
- configure Shape placement (`xlMoveAndSize`);
- detect closed/invalid documents;
- continue exporting into an existing document.

COM objects must remain inside the Excel infrastructure layer. Avalonia ViewModels and domain models must never hold raw Excel COM objects.

COM lifetime needs explicit handling. Avoid long chains of temporary COM property accesses and release owned RCWs deterministically where appropriate. Operations should execute on a controlled thread/apartment compatible with Office automation.

## 13. Progress and cancellation

Manual import and Excel export may process many images and therefore expose progress and cancellation.

Progress should report meaningful units, for example:

```text
Importing 23 / 120
Exporting 41 / 120
Normalizing worksheet
Preparing Excel
```

Cancellation should use `CancellationToken` through application services.

Cancellation is cooperative. Do not interrupt an Office COM call halfway through an unsafe mutation. Check cancellation between atomic/safe operations, such as before processing the next image.

The UI must remain responsive while import/export is running.

## 14. Configuration

Expected configuration areas include:

### Capture

- capture provider;
- monitored folder;
- monitored key(s);
- automatic mouse movement enabled;
- mouse X/Y;
- Greenshot executable/configuration location.

### Image sizing

- design width;
- design height;
- design scale;
- maximum enlargement;
- scale rule definitions;
- DPI behavior.

### Excel layout

- starting X/Y or cell/origin;
- horizontal image interval;
- vertical image interval;
- row height;
- column width;
- space-saving mode;
- template location;
- target worksheet behavior.

Configuration models should be versioned so future releases can migrate older settings.

## 15. Reliability considerations

Important failure cases include:

- Greenshot exits unexpectedly;
- Greenshot output file is observed before writing completes;
- duplicate `FileSystemWatcher` events;
- capture folder is missing or inaccessible;
- imported image is corrupt or unsupported;
- malformed scale rule;
- Excel is not installed;
- Excel startup fails;
- workbook/sheet is closed while RelaxCat still holds a reference;
- COM call fails because Excel is busy;
- user cancels during a multi-image operation;
- template is missing or invalid.

Failures affecting one image should not necessarily destroy the whole import/export session. Where safe, record an item-level error and continue or allow the user to retry.

## 16. Suggested implementation phases

### Phase 1 — Core image pipeline

Implement image metadata, DPI normalization, scale DSL, fallback design sizing, cache, row layout, and unit tests. No Excel or Greenshot dependency is needed to validate this layer.

### Phase 2 — Excel integration

Implement Excel session management, template creation, worksheet normalization, image insertion, `xlMoveAndSize`, workbook reuse, progress, cancellation, and recovery from closed documents.

### Phase 3 — Capture integration

Implement global key monitoring, mouse movement, Greenshot process lifecycle, folder monitoring/stability detection, and automatic import.

### Phase 4 — Avalonia UI

Complete configuration UI, image queue/preview, rule editor/validation, progress/cancel UI, status/error presentation, and capture-provider switching.

### Phase 5 — Additional capture providers and polish

Add/refine Windows native capture provider, diagnostics, configuration migration, packaging, and automated integration tests where practical.

## 17. Design principles

RelaxCat should favor:

- deterministic layout over clever packing;
- explicit platform boundaries over Windows calls scattered through UI code;
- small dedicated DSL over arbitrary script execution;
- configuration snapshots and pure calculations where possible;
- cache keys based on effective inputs;
- safe Office lifecycle management;
- cooperative cancellation;
- independently testable sizing/layout logic.

The core sizing and layout engine should be usable without Avalonia, Greenshot, or Excel. This is the main architectural improvement that allows RelaxCat to evolve beyond the original implementation without turning Windows/Office integration into the application architecture itself.
