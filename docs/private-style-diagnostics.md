# Local compatibility diagnostics

Use `RelaxCat.ExcelTool diagnostics workbook.xlsx --destructive-styles` to produce a sharing-oriented JSON summary. Omit the destructive option to diagnose conservative mode. This is read-only and does not optimize or upload the workbook. The desktop analysis page also offers a selectable JSON summary under its diagnostic summary expander.

The summary reports every **detected gate category**, not just the first reason. Categories represent presence, not counts of individual occurrences. Unsupported parts, unknown style extensions, package extensions and unsupported references are independent gates; clearing one does not imply the others have cleared. Failed parsing produces a generic CLI error without workbook details; a complete analysis cannot be promised for malformed input.

Retained-table categories are separate from blockers: phonetic references can pin the font table, and inline differential number formats can pin custom number-format IDs without preventing base-style cleanup. An empty blocker list only means the implemented scanner found no gate; it is not native Excel validation or a promise of visual equivalence in destructive mode.

Only fixed, allowlisted categories and aggregate format counts/sizes are exported. Paths, sheet names, style names, cell contents/formulas, external-link targets, exception messages and arbitrary XML namespace/extension strings are excluded. Aggregate counts and sizes may still be confidential; review before sharing. `analyze --json` is a detailed local report and is **not** this privacy-filtered export. The older `extensions` command also contains arbitrary namespace/type metadata and requires review.

Projected ZIP size is an estimate. Native Excel open/save/reopen compatibility and visual fidelity require separate validation on copies; this report does not claim either.

Ancillary parts with global number-format IDs are reported as `ancillaryNumberFormats`; the global number-format table remains pinned so those unmodified consumers keep their indices.
