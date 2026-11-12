# Non-empty workbook compatibility

A workbook's cell count is not a measure of style-table usage. A small sheet may
reference many cell formats, while tens of thousands of named styles may keep
base formats alive without any direct cell reference. Analyze both the worksheet
reference graph and the package compatibility report before drawing conclusions.
No provenance or intent (for example, deliberate pollution) can be inferred from
style counts alone.

## Supported ordinary parts

The advanced optimizer can now retain ordinary DrawingML drawings (including
shapes and text boxes), images, charts, tables, hyperlinks, comments and legacy
VML comment/header images, and basic pivot tables/caches. Their original part
bytes and relationships are preserved. The XML namespaces and style-reference
attributes are checked, rather than trusting a ZIP filename alone. Ordinary
images are never decoded or re-encoded, and external targets are never fetched.

Text-box and header/footer font names (such as `Microsoft YaHei`) are literal
font names, not indices in the workbook `fonts` table. They are preserved
verbatim. Worksheet/shared-string phonetic font indices still pin the complete
font table in destructive mode. External number-format references in supported
parts pin the complete `numFmts` table; differential formats (`dxfs`) are not
renumbered. Pinning a component table does not prevent named/base style cleanup.

Common Office metadata, including `x14ac:knownFonts`, `x14ac:dyDescent` and table
revision identifiers, is recognized in the explicit destructive compatibility
path. This is not blanket support for all Office extensions.

## Boundaries

Unknown XML extensions, signed packages, unsupported controls/OLE, unknown
relationships, and unmodeled style-name consumers still disable advanced edits.
Malformed auxiliary XML is reported as unsupported rather than being partly
rewritten. A normal feature may contain an unsupported extension, so the presence
of a picture or table alone is neither a failure nor a guarantee of support.

Safe mode keeps named-style identity and inheritance constraints. Explicit
`--destructive-styles` still discards named-style identities and future
inheritance, including those referenced by currently used cell formats, while
retaining explicit cell-format attributes and Normal's metrics. It does not
promise visual/semantic equivalence or VBA compatibility. Keep the original and
review an output copy. Macros and external links are not executed.

## Validation

`AncillaryStylePartTests` covers non-empty active style remapping, safe and
lossy advanced modes, preserved text-box/drawing/table/hyperlink bytes,
number-format pinning, and second-pass byte idempotence. The generated-workbook
matrix in `tools/verify-nonempty-style-matrix.py` exercises real OOXML emitted by
openpyxl with multiple sheets and active formats plus copied style inventories.
All fixtures are synthetic. No private workbook is needed or redistributed.

Package/reference checks and openpyxl reopen checks are not Microsoft Excel
rendering validation. Windows Excel open/save/reopen and visual comparison must
be reported separately when available, never inferred from these checks.
