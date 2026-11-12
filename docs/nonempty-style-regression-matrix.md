# Nonempty workbook regression matrix

`tools/verify-nonempty-style-matrix.py` creates synthetic workbooks from scratch with openpyxl and Pillow. It does not read issue attachments, production workbooks, account data or company documents.

## Running

```sh
python -m pip install openpyxl Pillow
dotnet build tools/RelaxCat.ExcelTool -c Release
python tools/verify-nonempty-style-matrix.py --output-dir /tmp/relaxcat-matrix \
  --tool 'dotnet tools/RelaxCat.ExcelTool/bin/Release/net10.0/RelaxCat.ExcelTool.dll'
```

The default is 48,846 base formats and named styles, with distinct identities but repeated base formatting. Six live cell formats refer to different copied base identities. `--styles 100` provides a fast smoke test. Omitting `--tool` validates only generated inputs, explicitly reported as such. Output is generated in the chosen directory and is not committed.

## Coverage and assertions

- Three sheets, including a hidden sheet, numeric and string values, 100 formulas, cross-sheet and structured-table formulas.
- Multiple live formats, custom number formats, row and column formatting, column width, frozen panes, merged cells, validation and differential conditional formatting.
- Worksheet DrawingML textbox/shape with explicit anchor, Microsoft YaHei Latin/East Asian font settings, size and bold. The entire drawing part must remain byte-identical.
- Real table, bar chart and PNG drawing objects; external hyperlink; comments and their VML drawing.
- Header with Microsoft YaHei font, explicit size and bold directives; footer page counters; synthetic header picture and VML/image relationships.
- A separate external-workbook link variant includes the complete relationship graph and cached external value.
- An unknown extension with a possible style reference must keep optimization blocked. This is a negative control, not a supported-extension claim.
- All non-style/non-worksheet parts must be byte-identical. Worksheet XML may change only cell, row and column style indices. Formulas, values, hyperlinks, relationships, table definitions, chart caches, pictures, headers and footers are preserved.
- Live style/component indices and differential-style references must remain in bounds; declared table counts must match; internal package relationships must resolve.
- Independently load both workbooks through openpyxl and compare cell values, formulas, direct formatting, merges, sheet state and loaded table/chart/image counts.
- Input file hash must not change. Running optimization twice must produce byte-identical packages.
- For the ordinary rich workbook, safe garbage collection must also succeed and retain named-style identities.
- Report both raw `styles.xml` bytes and actual ZIP entry compressed bytes, total package size, base/cell format counts and named styles in `matrix-report.json`.

## Verification limits

This test is package and independent object-model verification, not a claim of rendered equivalence or Microsoft Excel acceptance. Destructive cleanup deliberately discards named-style identity and inheritance behavior; the fixture checks supported direct-cell formatting, not arbitrary user editing or macros that depend on style names or indices.

openpyxl cannot calculate formulas and does not prove that a header font is installed. It does not render headers, VML or chart layouts. The generator's manually added header VML is a relationship-preservation probe, not evidence that Excel renders it identically. Native Excel 2007/2016/2024 open/save/reopen, repair warnings and visual comparison remain a separate Windows validation step. Do not save these workbooks through openpyxl as a substitute for that test; it can omit unsupported Office metadata.

Macro binary preservation and unknown/security-sensitive parts are tested separately in the .NET suite. No fake `vbaProject.bin` is represented as a valid executable VBA project by this matrix. Synthetic external links are never opened or fetched.

## Recorded synthetic run (2026-11-12)

.NET 10.0.404 CLI build succeeded with zero warnings/errors. The 100-style smoke matrix and full 48,846-style matrix both passed, including the rich-workbook safe-GC assertion. Full-run observations:

| Synthetic case | Package bytes, before → after | Raw styles.xml bytes | ZIP styles.xml bytes | Base/named styles |
| --- | ---: | ---: | ---: | ---: |
| Rich nonempty | 273,885 → 14,176 | 6,147,409 → 4,629 | 260,632 → 837 | 48,846 → 1 each |
| Rich + external workbook link | 274,847 → 15,148 | 6,147,409 → 4,629 | 260,632 → 837 | 48,846 → 1 each |
| Unknown extension negative control | 274,001 → 274,001 | 6,147,558 → 6,147,558 | 260,748 → 260,748 | 48,846 unchanged |

All three retained their six required cell formats, passed component-index and package-relationship checks, and produced byte-identical packages on the second optimization. The unknown-extension case was correctly blocked rather than reported as an optimization success. Safe GC retained all 48,846 named identities. Direct cell/row/column formatting, values, formulas and worksheet structure matched the independent loaded model; textbox, image, chart, VML, header text, table and external-link package bytes were preserved.

This run did **not** launch Microsoft Excel, render the worksheets or execute VBA. The measurements apply only to these generated inputs; they are not predictions for every workbook.
