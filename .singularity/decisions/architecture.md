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

Status: Superseded by D-PHASE-002

Upstream: S-IMPORT-001, S-SIZE-001, S-SIZE-002, S-SIZE-003, S-CACHE-001, S-LAYOUT-001, S-EXCEL-001, S-EXCEL-002, S-EXCEL-003, S-EXCEL-004, S-UI-001, S-I18N-001

Decision: The first runnable slice is Windows manual image import, queue/preview, editable sizing/rules/layout, and live Excel export. Capture providers, monitored folders, global hooks, provider packaging and durable settings are deferred. The initial technical-design phase order is superseded by this vertical slice; completed core logic is reused.

Reason: This gives a usable end-to-end workflow without presenting incomplete capture providers as working features. Deferred singularities remain authoritative for their future implementations.

Evidence: `src/RelaxCat.App/`, `README.md`, `docs/phase-one-validation.md`.

## D-UNITS-001

Status: Superseded by D-UNITS-002

Upstream: S-SIZE-001, S-SIZE-002, S-SIZE-003, S-LAYOUT-001

Decision: Keep existing Core display dimensions in 96-DPI logical pixels. Convert dimensions and design row width to Excel points with 72/96 at the export boundary. Spacing is configured in points. Rule scale is authoritative on a match; the maximum enlargement limit belongs to fallback only. Validate the complete rule text before modifying Excel. Parentheses remain unsupported; comparisons bind before `and`, then `or`.

Reason: Makes units explicit without rewriting the working sizing engine. The old rule-scale clamp contradicted S-SIZE-002 and was corrected, rather than promoted into intent.

Evidence: `src/RelaxCat.Core/Export/ImageExport.cs`, `tests/RelaxCat.Core.Tests/WorkflowTests.cs`.

## D-UI-001

Status: Superseded by D-UI-002

Upstream: S-UI-001, S-I18N-001, S-IMPORT-001

Decision: Use Avalonia with a dark, three-pane desktop-tool layout and an in-session Chinese/Japanese language switch. Maintain paired translations for application labels and status messages; append original exception details as diagnostics. Run image decoding off the UI thread, export on the COM STA worker, and cancel between files/shapes. Completed work remains when cancelled; under S-EXCEL-008, re-exporting clears pictures from the target worksheet and exports the selected queue again.

Reason: Implements the IDEA-style direction and both languages while keeping the first slice small. Explicit replacement/cancellation semantics avoid pretending Office mutations are transactional.

Evidence: Original `src/RelaxCat.App/MainWindow.cs` (replaced by `MainWindow.axaml` and `MainWindow.axaml.cs` under D-UI-002). The original full-queue export and three-pane arrangement are superseded by the human's S-UI-002 / S-IMPORT-002 instruction.

## D-UI-002

Status: Active

Upstream: S-UI-001, S-UI-002, S-UI-003, S-I18N-001, S-IMPORT-001, S-IMPORT-002, S-EXCEL-001

Decision: Declare the application theme resources, main shell/gallery/lightbox, and settings tabs in Avalonia XAML, with C# code-behind for platform interactions and observable models for state. The home page retains design scale and compact layout only. A separate settings view has sizing/rules, Excel layout, appearance/language, and capture/monitoring pages. Theme dictionaries supply light/dark semantic brushes; both themes and live language switching also apply to existing cards and preview. Capture actions stay visible below scrolling details.

Reason: Preserves the IDEA-style desktop-tool direction while following the human's explicit XAML, simpler-home and theme requirements. Shared setting bindings avoid duplicate values between home and settings.

Evidence: `src/RelaxCat.App/App.axaml`, `src/RelaxCat.App/MainWindow.axaml`, `src/RelaxCat.App/Views/SettingsView.axaml`, `src/RelaxCat.App/Models/`, `tests/RelaxCat.App.SmokeTests/Program.cs`.

## D-GALLERY-001

Status: Active

Upstream: S-IMPORT-001, S-IMPORT-002, S-SIZE-001

Decision: Present every queue item as a card in a scrollable wrapping gallery with a bounded thumbnail, original metadata, inclusion checkbox, copy action and remove action. Copy places full source-image pixels on the clipboard and flushes the Windows clipboard; remove releases the queue thumbnail without deleting the source file. New images are included by default. Export snapshots only checked items in queue order; zero selection does not create or modify Excel. Unchecked items remain available for preview and copying.

Reason: Gives ordinary meaning to copy/remove while keeping export participation independent of queue membership. Thumbnails are for UI only; export still embeds original files. Snapshotting and disabled mutation controls retain safe progress/cancellation boundaries.

Evidence: `src/RelaxCat.App/MainWindow.axaml.cs`, `src/RelaxCat.App/Models/ImageItem.cs`, `src/RelaxCat.App/Models/WorkspaceModel.cs`, `tests/RelaxCat.App.SmokeTests/Program.cs`.

## D-PREVIEW-001

Status: Active

Upstream: S-IMPORT-002, S-UI-001, S-UI-003, S-I18N-001

Decision: Open a themed in-window lightbox with dimmed backdrop, fade-in, metadata, shared inclusion checkbox, fit-to-window, zoom/scroll, previous/next controls and arrow/Esc shortcuts. Decode the preview on a worker with a bounded maximum dimension; discard results from an older preview request and release bitmaps on close. These preview-only reductions never alter source pixels or export sizing.

Reason: Provides a focused, responsive preview with predictable navigation and bounded per-image memory, including when users move rapidly between images.

Evidence: `src/RelaxCat.App/MainWindow.axaml`, `src/RelaxCat.App/MainWindow.axaml.cs`, `tests/RelaxCat.App.SmokeTests/Program.cs`.

## D-UNITS-002

Status: Active

Upstream: S-SIZE-001, S-SIZE-002, S-SIZE-003, S-SIZE-004, S-LAYOUT-001, S-LAYOUT-002, S-EXCEL-002

Decision: Retain 96-DPI logical display dimensions and the 72/96 Excel point conversion. Horizontal and vertical gaps independently choose points (existing default) or cell units, including fractions. After worksheet normalization, read one uniform cell's actual Width/Height in points through COM; convert horizontal cell gaps using Width and vertical gaps using Height. Do not approximate Excel character-based ColumnWidth as physical width. Use the configured vertical gap within a batch. S-EXCEL-008 supersedes appending: after validating layout and the start cell, clear target-sheet pictures and restart from the configured cell origin. Rule scaling, fallback limits and DSL precedence from D-UNITS-001 remain unchanged.

Reason: Cell geometry depends on Excel font/rounding and must be read from the normalized worksheet. Changing units must not change pixel sizing or export order.

Evidence: `src/RelaxCat.Core/Export/ImageExport.cs`, `src/RelaxCat.App/ExcelSession.cs`, `tests/RelaxCat.Core.Tests/ExportOptionsTests.cs`, `tests/RelaxCat.App.SmokeTests/ExcelChecks.cs`.

## D-ASPECT-001

Status: Active

Upstream: S-SIZE-001, S-SIZE-003, S-SIZE-004, S-EXCEL-002, S-EXCEL-003, S-EXCEL-004

Decision: Use the arithmetic mean of valid per-axis DPI (invalid axes fall back to 96) as a single uniform sizing DPI, retaining the original pixel aspect ratio for rule matches and fallback. Before RelaxCat normalizes worksheet rows/columns, snapshot each picture's geometry, aspect lock and cell placement, temporarily freeze it with xlFreeFloating, then restore geometry and original placement in a finally block. Newly inserted pictures explicitly receive the calculated width/height, LockAspectRatio=true, and xlMoveAndSize. Existing non-picture objects retain ordinary Excel normalization behavior.

Reason: Two actual defects contradicted the user's aspect-ratio requirement: independent axis DPI normalization warped non-square DPI metadata, and Excel resized existing xlMoveAndSize pictures non-uniformly when RelaxCat changed row/column sizes. Temporarily freezing pictures reconciles normalization with aspect preservation without weakening confirmed cell-following placement. This protects RelaxCat's own export operations; it does not retroactively reconstruct pictures already distorted by older versions or override independent user edits in Excel.

Evidence: New DPI regression tests first failed (a 1.5 pixel ratio became 3, 0.75 or 6.25). The live Excel regression first changed approximately 609.88×405.00 pt into 168.13×560.25 pt at 2.38-character columns; after the fix it remains 609.88×405.00 pt with xlMoveAndSize restored. See `tests/RelaxCat.Core.Tests/CoreTests.cs`, `tests/RelaxCat.App.SmokeTests/ExcelChecks.cs`.

## D-EXCEL-SELECTION-001

Status: Active

Upstream: S-EXCEL-005, S-IMPORT-002, S-IMPORT-001, S-I18N-001

Decision: Add an opt-in Excel settings checkbox, off by default. After successful nonempty export, select all picture and linked-picture Shapes on the target worksheet, including previous exports; exclude non-picture objects. Do not select on cancellation/failure and do not change the gallery's inclusion checkboxes. Use the Excel 2007-era Shapes.Range(...).Select API without grouping or modifying pictures.

Reason: In the supplied Excel-settings context, selecting pictures is an Excel editing convenience and is independent of deciding which queue images are exported. Opt-in preserves existing behavior.

Evidence: `src/RelaxCat.App/Views/SettingsView.axaml`, `src/RelaxCat.App/Models/SettingsModel.cs`, `src/RelaxCat.Core/Export/ImageExport.cs`, `src/RelaxCat.App/ExcelSession.cs`, `tests/RelaxCat.Core.Tests/ExportOptionsTests.cs`, `tests/RelaxCat.App.SmokeTests/ExcelChecks.cs`.

## D-PHASE-002

Status: Active

Upstream: S-CAPTURE-001, S-CAPTURE-002, S-CAPTURE-003, S-IMPORT-001, S-IMPORT-002, S-UI-002, S-I18N-001

Decision: Extend the existing runnable gallery/Excel slice with selectable Off, folder-only and bundled Greenshot modes, live folder-to-queue import, capture settings, owned-process lifecycle and optional cursor movement. Preserve existing sizing, checkbox filtering, preview and Excel adapters. Settings remain session-scoped and default to Off; applying a selected provider starts it. Direct Windows-native capture and durable configuration remain future work.

Reason: This fulfills the human's request to continue with Greenshot and monitoring without rebuilding finished functionality or inventing a new product intent.

Evidence: `src/RelaxCat.App/MainWindow.Capture.cs`, `src/RelaxCat.App/Models/CaptureSettings.cs`, `README.md`, `docs/capture-validation.md`.

## D-GREENSHOT-001

Status: Active

Upstream: S-CAPTURE-001, S-CAPTURE-003

Decision: Vendor the unmodified official full Windows portable pre-release 1.4.371-g78211b16d9, latest suitable upstream build checked on 2026-10-20. Preserve license/source links and the GitHub-published SHA-256. Copy the archive and notices into build/publish output; verify and extract into RelaxCat's own local-data cache on first use. Use --ini-directory with a process-scoped RelaxCat profile, PNG FileNoDialog output, unique capture filenames and update checking disabled. This full upstream package requires Windows x64 and .NET Framework 4.8.

Reason: Bundling avoids a machine-wide installation and makes the dependency reproducible/offline at runtime. The exact version/hash is a packaging decision, not a new singularity; future updates require reevaluating the pinned startup/configuration contract.

Evidence: `vendor/greenshot/dependency.json`, `vendor/greenshot/portable.zip`, `src/RelaxCat.App/Capture/GreenshotBundle.cs`, `src/RelaxCat.App/Capture/GreenshotProvider.cs`.

## D-GREENSHOT-LIFECYCLE-001

Status: Active

Upstream: S-CAPTURE-001, S-CAPTURE-003

Decision: Refuse to start if another Greenshot exists, because upstream uses a session-wide single-instance mutex. Manage only the Process that RelaxCat starts and its Job Object descendants. A kill-on-close job covers abnormal RelaxCat exit as well as explicit stop. Verify readiness by connecting to the SID-scoped pipe and checking that its server PID equals the owned PID, without issuing commands. Stop on provider switch, application close, startup cancellation and unexpected child exit. Do not broadcast Greenshot's global --exit or terminate by process name.

Reason: Real integration testing exposed that an early greenshot-cli probe can auto-start an extra instance. Passive PID-verified pipe readiness avoids that side effect, and process ownership prevents affecting unrelated user instances.

Evidence: `src/RelaxCat.App/Capture/GreenshotProvider.cs`, `src/RelaxCat.App/Capture/OwnedProcessJob.cs`, `tests/RelaxCat.App.SmokeTests/CaptureChecks.cs`; pinned upstream `GreenshotApplication.cs`, `GreenshotCommandLine.cs`, `NamedPipeEndpoint.cs`.

## D-CAPTURE-MONITOR-001

Status: Active

Upstream: S-IMPORT-001, S-CAPTURE-001, S-IMPORT-002

Decision: Watch only the selected folder's top level; baseline existing files at start. FileSystemWatcher notifications prompt rescans, backed by periodic full reconciliation for missed/overflowed events. Identify versions by full path, size and last-write time; wait 750 ms without changes and require a readable handle that excludes active writers. Import sequentially via the same ImportPathsAsync pipeline as manual import, awaiting busy manual/export operations. Retry decoding failures up to five times; report failures and retry a changed file version. Stop cancellation propagates to pending/active monitored imports without cancelling unrelated manual/export work.

Reason: Filename events alone do not imply complete image contents; notification coalescing and serial UI delivery avoid partial-file reads, duplicates and lost captures during export. Existing images can be manually imported instead of unexpectedly loading an entire folder.

Evidence: `src/RelaxCat.Core/Capture/CaptureFolderMonitor.cs`, `src/RelaxCat.App/MainWindow.Capture.cs`, `tests/RelaxCat.Core.Tests/CaptureTests.cs`, `tests/RelaxCat.App.SmokeTests/CaptureChecks.cs`.

## D-CAPTURE-KEYBOARD-001

Status: Retired

Upstream: S-CAPTURE-002, S-CAPTURE-001

Decision: When cursor movement is enabled for Greenshot, install a WH_KEYBOARD_LL hook on the UI message-loop thread after the provider is ready. Parse configured virtual keys/modifiers into an immutable options snapshot; on matching key-down synchronously call SetCursorPos, then always CallNextHookEx with the unchanged event, including error cases. Keep heavy work out of the callback. Unhook before stopping/switching providers. Configured coordinates may be negative for virtual desktops.

Reason: The same key must reach Greenshot after movement, so neither RegisterHotKey interception nor asynchronously moving the pointer is suitable. Latest-hook-first ordering and the short synchronous callback implement the required sequence.

Evidence: `src/RelaxCat.App/Capture/KeyboardMouseMonitor.cs`, `src/RelaxCat.Core/Capture/CaptureShortcut.cs`, `tests/RelaxCat.App.SmokeTests/CaptureChecks.cs`. A real F24 key test moved the cursor and still produced a Greenshot capture that reached the gallery.

Retirement: S-CAPTURE-002 was retired with human authorization after the combined screenshot/cursor verification failed. The hook and UI have been removed; Greenshot continues to own capture hotkeys.

## D-WORKSPACE-ORDER-001

Status: Active

Upstream: S-IMPORT-003, S-IMPORT-004, S-WORKFLOW-001

Decision: Snapshot file last-write UTC at import, default to oldest first, and support newest first and original import order. Use a monotonic per-item import sequence for ties. Insert new items into the current order and move existing item objects when sorting, preserving selection and preview identity. Export uses stable image IDs for source mapping. Clear-all releases item resources and clears the collection without deleting files or touching Excel; it is disabled while busy and does not stop monitoring.

Reason: File last-write time is available consistently for screenshots and manual imports; UTC comparisons avoid local-time ambiguity. Retaining item objects keeps per-image state intact. No export tracking or automatic export is introduced.
