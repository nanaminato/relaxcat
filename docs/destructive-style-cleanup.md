# Destructive style cleanup: independent experimental preview

**This is a lossy mode, not safe deduplication. Keep the original and use disposable copies.**

The branch `experiment/destructive-style-collapse` starts from the newest implementation in PR #2, `75c42cb426337c38cd6bbd67fc3698e563256333`. The default branch `main` was older (`d92b6150deadba551e0f7d48fea80900844bf4bb`); neither `main` nor `master` is modified or merged. The preview is independent of the existing releases.

## What changes

Enable the separate default-off destructive checkbox on the single-file or batch screen, or pass `--destructive-styles` to CLI `analyze`, `optimize` or `batch`. It forces unused-definition collection and component merging even without `--gc`. The default optimizer remains conservative.

1. Keep exactly one built-in Normal identity, recognized by `builtinId=0`, including its existing localized name and base definition. Move it to base index 0. If none exists, create Normal using old base 0, or an empty base if the base table was absent.
2. Delete every other named identity and base definition, including built-ins, hidden styles, metadata-bearing styles and names still associated with cells. There is no suffix heuristic and no promise that those names were unused.
3. Keep each cell XF's direct attributes and children, set its `xfId=0`, and set all six `apply*` flags to 1, intentionally breaking future inheritance/propagation. Do **not** replace direct formatting with base values when the old apply flag was false. Missing direct components stay missing; the tool does not guess inheritance. Keep the original Normal base/font because it can affect column-width metrics.
4. Collect unreferenced cell XFs, then unused components; merge exact component and normalized cell-format duplicates. Keep reserved default entries and remap known cell, row, column and style-table indices. Differential formats and other package parts are retained.
5. Verify output against the selected lossy plan, valid known indices, unique named ownership, worksheet counts/merges and exact expected package bytes. This does **not** verify equivalence to the original style semantics or an Excel-open guarantee.

`IsLossy=true`, `VerificationScope=KnownStructureAndContentOnly_NotVisualOrSemanticEquivalence`, `Warnings` and separate `DestructiveRemoved` counts are exposed in API/JSON. `Verified=true` means only the stated scope passed. Removed identity mappings use -1, even though cells formerly associated with those identities are rebound to Normal. Original paths are never overwritten. A subsequent unchanged destructive pass copies the ZIP verbatim.

## Intentional loss and remaining structural blocks

- Named-style editing independence and future inheritance are gone. Applying a deleted style, template workflows, macros using `Styles(name/index)` or `Range.Style`, and external automation may fail or behave differently.
- Current formatting, protection, quote handling, printing and later editing behavior are not guaranteed across Excel versions. Removing built-ins may change the style gallery or cause Excel to recreate definitions on save.
- VBA is admitted as an explicit behavior risk only in this mode; binary bytes are retained, not macro semantics. Tests use a synthetic binary sentinel, not a real Office macro execution test.
- Unknown XML extensions, unknown package parts and digital-signature parts still block index changes. For example, ordinary drawings/tables outside the current reference audit can still produce `unsupportedParts`. This preview does not indiscriminately drop workbook features or leave unknown index consumers pointing to removed entries. The skipped reason remains visible and no style reduction is claimed for a blocked plan.
- Invalid known component/base/cell references still fail rather than producing corrupt output. Duplicate named identities may be discarded in this mode, but the output must have valid unique ownership.
- The existing inline-DXF/custom-number-format gate remains. This is an aggressive named/base cleanup, not a claim to minimize every possible OOXML feature.

## Reproducible synthetic benchmark

No company file or TestHelper was opened, copied, uploaded or used. The fixture is generated entirely by `tools/generate-destructive-style-benchmark.py`.

```sh
python tools/generate-destructive-style-benchmark.py --output synthetic.xlsx
# SDK 10.0.300 as pinned in global.json
# Use a NEW path each time; keep the original.
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- analyze synthetic.xlsx --destructive-styles --json
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize synthetic.xlsx --output destructive.xlsx --destructive-styles --json
dotnet run --project tools/RelaxCat.ExcelTool -c Release -- optimize destructive.xlsx --output second.xlsx --destructive-styles --json
```

Fixture: 48,846 named identities, 48,846 bases and 48,846 cell XFs in 24 direct-format families; 12,000 rows, 59,996 materialized cells, 11,999 formulas with cached values, row/column styles, a merge, conditional formatting, validation, freeze pane, defined name and core metadata.

Measured on Linux with .NET SDK 10.0.300:

| Metric | Original | First destructive output |
| --- | ---: | ---: |
| Named styles | 48,846 | 1 |
| Base XFs (`cellStyleXfs`) | 48,846 | 1 |
| Cell XFs (`cellXfs`) | 48,846 | 24 |
| `styles.xml`, uncompressed bytes | 14,608,104 | 6,012 |
| `styles.xml`, ZIP-compressed bytes | 506,661 | 684 |
| Whole XLSX bytes | 1,028,494 | 513,674 |

Uncompressed style data reduction: **99.9588%**. Whole-package reduction: about **50.05%**. This is real removal of style XML, not just ZIP recompression. 71,990 explicit worksheet references remapped. One first-pass timing was 9.22 seconds including CLI startup; this is a single observation, not a performance guarantee.

Second pass: no table changes or remapped references; style XML and full ZIP bytes identical to the first output after the no-op-copy fix. Original SHA-256: `85da06555725d0057c0bd2f95a719c7b7e67f0f79b9267c8f276a95e8bcd4fc9`. First/second output SHA-256: `53348b49aac6b016a1a9b1ec23821a52e4cb34c2fded537d53a2d43ce57b591b`.

## Verification performed

- Full solution Release build with exact SDK 10.0.300: 0 warnings/errors. `-m:1` avoids environment-specific parallel MSBuild named-pipe permission failures. The final solution build and test use this serial setting.
- Core regression suite: **153/153 passed**, including 16 new destructive cases. Covers 48,846 identities, direct apply=false values, localized Normal at a nonzero base with nonzero font, absent optional base/component tables, duplicate owners/names, formulas and cached values, byte-preserving row/column remaps, source protection, full-package second-pass idempotence, VBA sentinel preservation, unknown extension/signature blocking and invalid-index rejection.
- App model-only smoke: default-off single/batch options, four-language warnings, analysis invalidation, loss scope, source bytes and removal counts when Normal or all names were absent.
- Four-language validation: 306 keys each, matching placeholders and referenced keys.
- CLI analyze/optimize/batch JSON: parses, reports loss/scope/warnings, warnings on stderr, input bytes unchanged.
- Independent read-only review found an absent-component fallback exception and misleading GUI count/file-validity wording. All were fixed and regression-covered/rechecked.
- Independent reader `openpyxl 3.1.5`: opened before/after; compared 60,000 rectangular positions for values/formulas and direct font/fill/border/alignment/protection/number format; merges, freeze pane and validation retained. Non-style package parts are byte-identical; worksheet bytes differ only at intended numeric style references. This is not an Office compatibility substitute.

## Not verified

Microsoft Excel 2007/2016/2024 is unavailable here. Actual open/save/reopen without repair, visual rendering, printing, native protection semantics, macro execution and named-style behavior after future edits are **unverified**. Windows portable outputs are cross-built and ZIP/contents checked, but not launched on Windows. Full rendered GUI smoke is not a passing test in this Linux font environment. No production or company document was tested. Use the preview only if these risks are acceptable; do not replace a company original with its output.
