# Private unsupported-part diagnostics

The GUI's existing copyable diagnostic JSON and the CLI `diagnostics` command now use schema version 2. `UnsupportedPartDetails` explains the previously undifferentiated `unsupportedParts` gate without exporting document metadata. No document upload is needed.

```sh
RelaxCat.ExcelTool diagnostics input.xlsx --destructive-styles
```

Example (synthetic only):

```json
"UnsupportedPartDetails": [
  { "PartKind": "Drawing", "Reason": "ExtensionList", "PartCount": 1 },
  { "PartKind": "VmlDrawing", "Reason": "VmlNonCommentObject", "PartCount": 1 }
]
```

Each rejected package part contributes **one first rejection**, grouped by fixed category and reason. A part may contain additional unsupported structures; the report is not a complete defect inventory. No path, sheet/style/object name, text, formula, external target, namespace URI or arbitrary XML is emitted. Aggregate counts and sizes may themselves be confidential. Do not share the full `analyze` output instead of this projection.

## Interpreting the fixed reasons

- `UnrecognizedPart`: a part outside the recognized package patterns. `PartKind` narrows it to drawing, media, printer settings, external link, control, embedded object, signature, custom XML, etc.; `Other` covers anything else.
- `MalformedXml`, `UnexpectedRoot`: XML cannot be parsed safely or its root does not match the recognized part.
- `ExtensionList`, `AlternateContent`: extension or compatibility branch whose style dependencies are not modeled.
- `UnknownElementNamespace`, `UnknownAttributeNamespace`: unexpected namespace. Its spelling is deliberately never reported.
- `UnhandledStyleReference`, `InvalidStyleReference`, `NamedStyleReference`: an unsupported style consumer or invalid index. A table or pivot may consume a named style that destructive removal would otherwise delete.
- `VmlNonCommentObject`, `VmlUnsupportedObject`: VML includes something beyond supported legacy notes, such as controls or text-box client data.
- `RelationshipOwner`, `RelationshipType`, `InvalidRelationship`: an unrecognized relationship owner/type, missing target or malformed relationship. The tool never follows external targets.

`phoneticFonts` is a retained table, **not** a remapping blocker. It keeps font indices stable while other tables can be optimized, provided no blocker remains.

## Narrow additional compatibility

A standard DrawingML shape's `cNvPr` creation identity is admitted only when its extension list contains exactly one `a16:creationId`, the exact Office extension identifier, a valid braced GUID, and no extra attributes, children or non-whitespace text. It is object identity metadata, not a workbook style index. Original drawing bytes are preserved. Unknown extensions, mixed extension lists, other parents and additional reference-like attributes remain blocked. This does not claim that creation identity is the cause in any particular user's workbook.

## Validation

Synthetic tests cover all 15 fixed rejection reasons, aggregate repeated parts, prevent private markers from appearing in JSON, preserve blocked styles and auxiliary bytes, and exercise valid creation identity plus seven rejected mutations. Both supported and rejected creation-identity cases are byte-idempotent on a second optimization. GUI model tests check that the same fixed diagnostics reach all four language modes and are cleared on input/mode changes.

Native Microsoft Excel opening/saving and a real Windows GUI have not been exercised in the Linux test environment. A remaining blocker is reported rather than bypassed.

### Synthetic large-table check

A generated workbook containing 48,847 base/named styles and a shape creation-identity extension was optimized by the CLI: base styles 48,847 → 1; named styles 48,847 → 1; styles XML 5,549,299 → 3,206 bytes; ZIP-compressed styles 257,609 → 689 bytes; whole package 262,249 → 5,334 bytes. These are synthetic fixture measurements, not a promise about an unexamined document. Unknown-extension fixtures still report their fixed rejection reason and preserve the original style bytes.
