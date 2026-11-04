# Style graph regression validation

The `GraphStyleTests` suite creates synthetic OOXML packages in a per-test temporary directory and deletes them after the test. No user workbook, Excel installation, or network connection is required at test runtime.

Run:

```sh
dotnet test tests/RelaxCat.Core.Tests/RelaxCat.Core.Tests.csproj --filter FullyQualifiedName~GraphStyleTests
```

## Coverage

- 48,846 base styles and 48,846 separately named styles. Boolean lexical variants (`true`/`1`) yield two semantic base definitions and two cell formats. Every named-style identity is retained, and every original base/cell index is checked against the returned map.
- Exact versus semantic duplicate diagnostics, safe removal counts, graph completeness, and edges from named styles, cell formats, and worksheet style references.
- Original source bytes, formula text and cached values, inline strings, merged ranges, row/column inheritance, effective formatting previews, and all unchanged OOXML part payloads.
- Separate uncompressed XML, compressed ZIP-entry, and total workbook size measurements. A reduction in XML size is not treated as an equivalent reduction in workbook size.
- A second optimization is payload-idempotent: every ZIP entry has the same uncompressed bytes, and no worksheet references need remapping. ZIP container byte identity is not required.
- Missing apply flags are not conflated with explicit false values. Distinct alignment and protection settings remain distinct.
- Unknown worksheet extension references prevent all index-changing optimization, both with and without optional garbage collection. A macro-bearing package also preserves every index and the binary payload.
- Garbage collection remains opt-in. Named styles keep otherwise unused base styles and their components reachable.
- Explicit copied-name cleanup removes only the unambiguous metadata-free `Custom_1` candidate with an existing equivalent `Custom` style. Default behavior preserves every named style. Differing formats, built-in/custom-built-in markers, hidden/level metadata, unrelated names, leading-zero suffixes, nested suffixes, and case-insensitive ambiguities are retained. Named maps, GC composition, and second-pass idempotence are checked.
- Cleanup is a name-based heuristic, not proof that an external consumer does not use that name. External programs or templates may depend on copied-looking names; the option remains off by default and must only be enabled after considering that risk.
- Malformed explicit cell/row style indices and overflowing column indices are rejected even for otherwise unchanged packages and packages blocked by unknown extensions.
- Lexical `false`/`0` apply flags give the same inherited bold font and right alignment when the cell's own font and alignment deliberately differ.
- 20,000 repeated cell formats collapse to one, with every original index checked against the returned map.
- An attribute-looking substring in a named-style name (`A xfId='1'`) is preserved literally while the real `xfId` is remapped.

The wider `ExcelStyleComponentTests` and `ExcelStyleTests` suites cover component deduplication, reserved fill slots, sparse/custom number format IDs, invalid references, multiple XML encodings, binary/macro parts, cancellation, and unsupported package parts.

These tests verify package structure, preserved payloads and supported formatting semantics. They do not replace opening the output in supported Excel versions to validate application-specific rendering and repair behavior.

## Recorded synthetic run

On 2026-11-04, the full Release core test suite passed **123/123** tests using .NET SDK 10.0.403. The large graph case produced:

| Measurement | Before | After |
| --- | ---: | ---: |
| Base style definitions | 48,846 | 2 |
| Cell format definitions | 3 | 2 |
| Named style definitions | 48,846 | 48,846 |
| Uncompressed styles XML bytes | 5,766,983 | 2,187,980 |
| Compressed styles ZIP entry bytes | 435,968 | 129,127 |
| Complete workbook bytes | 438,688 | 131,498 |

These measurements describe this generated fixture, not expected savings for an arbitrary workbook. Runtime/ZIP sizes can vary with the SDK and compression implementation.

## Conservative versus explicit copied-name cleanup

The default retains every named style. `--remove-copied-named-styles` is a separate CLI opt-in (also `ExcelStyleOptimizationOptions.RemoveCopiedNamedStyles` for callers), not the existing `--gc` option. Run `analyze --json --remove-copied-named-styles` to review `NamedStyleCleanupCandidates` before optimizing.

Only an unambiguous custom `Name_N` (positive integer suffix) with an existing unsuffixed custom `Name` and equivalent base formatting is a candidate. Both entries must have only `name` and `xfId` metadata. Built-ins, custom built-ins, hidden styles, level metadata, unknown metadata, ambiguous duplicate names and nested suffixes are retained. No VBA/unknown-reference package is altered. `NamedStyleMap` records each retained index and `-1` for an explicitly requested name removal. `ConditionalNamedStyleRemovalCount` is separate from safe duplicate removal counts.

**Risk:** a naming pattern does not prove that a style is unused. An external program or template can refer to a name that is not referenced anywhere in the OOXML package. The CLI prints this warning whenever the option is supplied. Keep the source workbook and review candidates. There is deliberately no automatic deletion by name in the default path. The desktop single-file and batch pages also expose default-off controls with localized risk guidance.

The graph covers the remappable `cellStyles → cellStyleXfs ← cellXfs ← worksheets` chain and global components. Differential formats remain untouched and are not advertised as part of this graph. For `numFmts`, `TargetIndex` is the custom `numFmtId`, not the ordinal; built-in number formats use the separate `builtinNumberFormats` target.

## Environment and remaining validation

- Full solution Release build: passed with zero warnings/errors using SDK 10.0.403. The repository pins 10.0.300; only the local test run temporarily allowed `latestFeature` roll-forward, and the committed pin is unchanged.
- Four-language key/placeholder validation: passed. Desktop cleanup controls and smoke-test assertions compile.
- CLI `analyze --json --remove-copied-named-styles` and `optimize --remove-copied-named-styles --gc`: exercised end-to-end on a separate synthetic `Revenue`/`Revenue_1` fixture. Candidate reporting, warning emission and output verification passed.
- Desktop runtime smoke was attempted but stopped in the pre-existing localization window test before the Excel checks: this Linux environment cannot create the Windows `Segoe UI` composite glyph typeface. This is a blocked runtime check, not a passing UI test.
- Microsoft Excel is not available here. No real Excel open/save/reopen test or Excel 2007/2016/2024 rendering/repair claim is made. All workbooks used were generated test fixtures, not company documents.

## Local extension type diagnosis

The analysis UI now includes bounded extension type metadata. It lists extension URI,
namespace URI, element/attribute local names and occurrence counts, never ordinary
attribute values or text. Namespace declarations and foreign attributes outside
`extLst` are included as well. The scope is styles, workbook and worksheet XML,
not every package part. The diagnostic is not a reference graph or safety proof.
Unknown extensions continue to block all style-index changes, including GC and
optional copied-name cleanup.

For the smallest reviewable output, use the CLI:

```
dotnet RelaxCat.ExcelTool.dll extensions workbook.xlsx
```

This command emits only the type diagnostic JSON plus a fixed warning/scope. Errors
are generic and do not include paths or document values. Do not share `analyze --json`
or whole-app screenshots expecting the same restriction: those existing views include
file paths, worksheet names and potentially named-style cleanup candidates.

Nothing is uploaded by this feature. Review locally before sharing even this reduced
output: a custom URI, namespace or XML name can itself disclose private information.
There are at most 64 entries and 128 characters per displayed name. Control, bidi and
line/paragraph separators are replaced. `Truncated` warns when information was omitted
or shortened; collisions between shortened names can aggregate counts. Counts are
metadata occurrences, not style usage or known references. The UI uses localized labels
and warnings in Chinese, Japanese, English and Korean.

Microsoft's [MS-XLSX Styles extension table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d8afb021-974d-403a-84de-21af334d3d1d)
lists slicerStyles, timelineStyles, dxfs and xfComplementExt among the possible types.
These examples do not identify the user's actual extensions and are not an allowlist.
For example, [timeline style elements can reference dxfId](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/79d1e43b-5b1a-49e1-957c-50b49ff9a2e8).
Synthetic tests cover hidden text/value sentinels, foreign attributes without extLst,
control characters, bounded output and continued safety blocking.

Extension diagnostic validation checkpoint: full Release solution build passed with zero
warnings/errors; core suite 128/128 passed; four language packs 299 keys each passed;
`--excel-diagnostics-model-only` passed localization, bounded metadata and unchanged
safety-block checks without requiring desktop fonts. The dedicated CLI command passed
an independent JSON schema check and private-path error sentinel check. Full rendered
UI smoke and real Excel round trips remain unverified for the limits above.
