# RelaxCat Derived Architecture Decisions

These are agent-revisable implementation decisions derived from confirmed product intent. They are not human intent by themselves.

## D-ARCH-001

Status: Active

Upstream: S-CAPTURE-001, S-CAPTURE-002, S-SIZE-001, S-EXCEL-001

Decision: Keep image sizing/layout logic independent from Avalonia, Greenshot, Win32 hooks, and Excel COM. Platform integrations are adapters around a testable core.

Reason: The confirmed behavior does not require UI or platform APIs to participate in pure sizing/layout calculations, and separating them reduces coupling and makes agent changes easier to verify.

Evidence: `docs/technical-design.md`, `src/RelaxCat.Core/`.

## D-RULE-001

Status: Active

Upstream: S-SIZE-002

Decision: Implement the scale language as a small deterministic parser/evaluator rather than executing arbitrary C# or another general scripting language.

Reason: The confirmed language is intentionally small and ordered. A dedicated evaluator is safer, easier to cache, and easier to test. Grammar details not explicitly confirmed may evolve without changing S-SIZE-002.

Evidence: `src/RelaxCat.Core/Imaging/ScaleRuleEngine.cs`.

## D-LAYOUT-001

Status: Active

Upstream: S-LAYOUT-001

Decision: Use deterministic shelf/row packing for space-saving layout rather than general bin packing.

Reason: It directly implements the confirmed left-to-right, then next-row behavior and keeps Excel output predictable.

Evidence: `src/RelaxCat.Core/Layout/ImageLayoutEngine.cs`.

## D-CACHE-001

Status: Active

Upstream: S-CACHE-001, S-SIZE-001, S-SIZE-002, S-SIZE-003

Decision: Cache scale results by effective image metadata plus effective sizing configuration/rule identity. Cache invalidation follows key changes rather than treating cached values as durable intent.

Reason: Any input that can alter scale must distinguish cache entries.

Evidence: `src/RelaxCat.Core/Imaging/ScaleCache.cs`.

## D-EXCEL-001

Status: Active

Upstream: S-EXCEL-001, S-EXCEL-002, S-EXCEL-003, S-EXCEL-004

Decision: Implement late-bound Office COM automation on a dedicated STA worker, using the Excel 2007-era Workbooks, Worksheets and Shapes APIs without a version-specific interop assembly. Keep raw COM objects out of UI and Core. Reuse the workbook owned by this RelaxCat session; recover from closed workbooks with the configured template (blank means Excel's default template). Leave visible workbooks open for the user to save on exit.

Reason: The required live workbook reuse, worksheet normalization, shape insertion, and `xlMoveAndSize` behavior fit Office automation, while isolation contains COM lifecycle concerns.

Evidence: `src/RelaxCat.App/ExcelSession.cs`, `tests/RelaxCat.App.SmokeTests/ExcelChecks.cs`. This revises the old Excel-2016-specific proposal to comply with S-EXCEL-004; it does not change any confirmed intent. Version-specific manual validation is still required.

## D-PHASE-001

Status: Active

Upstream: S-IMPORT-001, S-SIZE-001, S-SIZE-002, S-SIZE-003, S-CACHE-001, S-LAYOUT-001, S-EXCEL-001, S-EXCEL-002, S-EXCEL-003, S-EXCEL-004, S-UI-001, S-I18N-001

Decision: The first runnable slice is Windows manual image import, queue/preview, editable sizing/rules/layout, and live Excel export. Capture providers, monitored folders, global hooks, provider packaging and durable settings are deferred. The initial technical-design phase order is superseded by this vertical slice; completed core logic is reused.

Reason: This gives a usable end-to-end workflow without presenting incomplete capture providers as working features. Deferred singularities remain authoritative for their future implementations.

Evidence: `src/RelaxCat.App/`, `README.md`, `docs/phase-one-validation.md`.

## D-UNITS-001

Status: Active

Upstream: S-SIZE-001, S-SIZE-002, S-SIZE-003, S-LAYOUT-001

Decision: Keep existing Core display dimensions in 96-DPI logical pixels. Convert dimensions and design row width to Excel points with 72/96 at the export boundary. Spacing is configured in points. Rule scale is authoritative on a match; the maximum enlargement limit belongs to fallback only. Validate the complete rule text before modifying Excel. Parentheses remain unsupported; comparisons bind before `and`, then `or`.

Reason: Makes units explicit without rewriting the working sizing engine. The old rule-scale clamp contradicted S-SIZE-002 and was corrected, rather than promoted into intent.

Evidence: `src/RelaxCat.Core/Export/ImageExport.cs`, `tests/RelaxCat.Core.Tests/WorkflowTests.cs`.

## D-UI-001

Status: Active

Upstream: S-UI-001, S-I18N-001, S-IMPORT-001

Decision: Use Avalonia with a dark, three-pane desktop-tool layout and an in-session Chinese/Japanese language switch. Maintain paired translations for application labels and status messages; append original exception details as diagnostics. Run image decoding off the UI thread, export on the COM STA worker, and cancel between files/shapes. Completed work remains when cancelled; re-exporting appends the current queue again.

Reason: Implements the IDEA-style direction and both languages while keeping the first slice small. Explicit append/cancellation semantics avoid pretending Office mutations are transactional.

Evidence: `src/RelaxCat.App/MainWindow.cs`.
