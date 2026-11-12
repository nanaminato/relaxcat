using System.Xml;

namespace RelaxCat.Core.Excel;

/// <summary>Conservative closed-world gate for table GC/component index changes.</summary>
internal sealed record StyleOptimizationContext(HashSet<int> CellReferences, string? SkipReason)
{
    internal IReadOnlyList<StyleReferenceEdge> AdditionalReferences { get; init; } = [];
    internal IReadOnlyList<string> SkipReasons { get; init; } = [];
    internal bool PreserveNumberFormatTable => AdditionalReferences.Any(edge => edge.Target == "numFmts");
    internal bool PreserveFontTable => AdditionalReferences.Any(edge => edge.Target == "fonts");

    internal static StyleOptimizationContext Read(OoxmlPackage package, StyleSheetDocument styles,
        IReadOnlyList<WorkbookSheet> worksheets, CancellationToken token, HashSet<int>? knownReferences = null,
        bool destructiveStyleCleanup = false)
    {
        var references = knownReferences ?? new HashSet<int>();
        var additionalReferences = new List<StyleReferenceEdge>();
        references.Add(0); // Excel's implicit/default cell format is always a root.
        var reasons = new List<string>();
        void AddReason(string value) { if (!reasons.Contains(value)) reasons.Add(value); }
        if (!(StyleDeduplicationPlan.CanMergeBaseStyles(styles) ||
              (destructiveStyleCleanup && KnownOfficeMetadata.CanRemapStyles(styles)))) AddReason("styleExtensions");
        foreach (var element in styles.Root.Descendants())
        {
            token.ThrowIfCancellationRequested();
            if (destructiveStyleCleanup && KnownOfficeMetadata.IsStyleExtension(element)) continue;
            var parent = element.Parent?.Name.LocalName;
            if (!KnownFormat(element, styles.Root.Name.NamespaceName, destructiveStyleCleanup)) AddReason("unsupportedReferences");
            var xf = element.Name.LocalName == "xf" && parent is "cellXfs" or "cellStyleXfs";
            foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            {
                var allowed = attribute.Name.LocalName switch
                {
                    "fontId" or "fillId" or "borderId" => xf,
                    "xfId" => (element.Name.LocalName == "xf" && parent == "cellXfs") ||
                              (element.Name.LocalName == "cellStyle" && parent == "cellStyles"),
                    "numFmtId" => xf || (element.Name.LocalName == "numFmt" &&
                        (parent == "numFmts" || element.Ancestors().Any(a => a.Name.LocalName == "dxfs"))),
                    _ => true
                };
                if (!allowed) AddReason("unsupportedReferences");
            }
        }
        var parts = new HashSet<string>(worksheets.Select(s => s.PartName), StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
            "xl/styles.xml", "xl/sharedStrings.xml", "xl/calcChain.xml", "xl/theme/theme1.xml",
            "docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"
        };
        // Explicit lossy policy may break macro name/index assumptions. Preserve VBA bytes,
        // but do not allow signed packages, unknown XML consumers or controls. Only
        // explicitly recognized metadata/ancillary parts are accepted below.
        if (destructiveStyleCleanup) parts.Add("xl/vbaProject.bin");
        foreach (var entry in package.Archive.Entries)
        {
            token.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith('/')) continue;
            if (!parts.Contains(entry.FullName) &&
                !(destructiveStyleCleanup && KnownOfficeMetadata.IsAncillaryPart(package, entry.FullName, worksheets)) &&
                !AncillaryStyleParts.Accept(package, entry.FullName, styles, additionalReferences, token)) AddReason("unsupportedParts");
        }
        foreach (var sheet in worksheets)
        {
            token.ThrowIfCancellationRequested();
            if (knownReferences is null)
            {
                using var source = package.OpenPart(sheet.PartName);
                var scan = WorksheetStyleScanner.Scan(source, styles.CellXfs.Count, sheet.Name, token);
                references.UnionWith(scan.RawStyleCounts.Keys);
                references.UnionWith(scan.EffectiveStyleCounts.Keys);
            }
        }
        // Do not inspect values or retain worksheet XML. Reject extensions and reference-like
        // attributes outside the explicitly supported graph, including unknown workbook metadata.
        foreach (var name in parts.Where(n => n.StartsWith("xl/", StringComparison.Ordinal) &&
                     n.EndsWith(".xml", StringComparison.Ordinal) && n != "xl/styles.xml" &&
                     !n.StartsWith("xl/theme/", StringComparison.Ordinal) && package.Find(n) is not null))
        {
            using var source = package.OpenPart(name);
            using var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            var ordinal = -1;
            while (reader.Read())
            {
                token.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element) continue;
                ordinal++;
                if (reader.NamespaceURI != styles.Root.Name.NamespaceName || reader.LocalName == "extLst")
                {
                    if (destructiveStyleCleanup && KnownOfficeMetadata.ReadPackageElement(reader, name)) continue;
                    AddReason("packageExtensions");
                    continue;
                }
                var elementName = reader.LocalName;
                var phoneticFont = destructiveStyleCleanup && elementName == "phoneticPr" &&
                    (name == "xl/sharedStrings.xml" || worksheets.Any(sheet => sheet.PartName == name));
                if (phoneticFont)
                {
                    if (!int.TryParse(reader.GetAttribute("fontId"), out var fontId) || fontId < 0 || fontId >= styles.Fonts.Count)
                        throw new InvalidDataException("Invalid phonetic font reference.");
                    additionalReferences.Add(new StyleReferenceEdge(name, ordinal, "fonts", fontId, "fontId"));
                }
                while (reader.MoveToNextAttribute())
                {
                    if (reader.NamespaceURI.Length > 0 && reader.NamespaceURI is not
                        ("http://www.w3.org/2000/xmlns/" or "http://www.w3.org/XML/1998/namespace" or
                         "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
                         "http://purl.oclc.org/ooxml/officeDocument/relationships") &&
                        !(destructiveStyleCleanup && KnownOfficeMetadata.IsAttribute(elementName, reader.NamespaceURI, reader.LocalName)))
                        AddReason("packageExtensions");
                    if ((reader.LocalName is "fontId" or "fillId" or "borderId" or "numFmtId" or "xfId" &&
                         !(phoneticFont && reader.LocalName == "fontId" && reader.NamespaceURI.Length == 0)) ||
                        (reader.LocalName == "s" && elementName is not ("c" or "row")) ||
                        (reader.LocalName == "style" && elementName != "col"))
                        AddReason("unsupportedReferences");
                }
                reader.MoveToElement();
            }
        }
        return new StyleOptimizationContext(references, reasons.FirstOrDefault()) { AdditionalReferences = additionalReferences, SkipReasons = reasons };
    }
    private static readonly HashSet<string> KnownStyleElements = new(
        ("numFmts numFmt fonts font name charset family b i strike outline shadow condense extend color sz u vertAlign scheme " +
         "fills fill patternFill fgColor bgColor gradientFill stop borders border left right top bottom diagonal vertical horizontal start end " +
         "cellStyleXfs cellXfs xf alignment protection cellStyles cellStyle dxfs dxf tableStyles tableStyle tableStyleElement " +
         "colors indexedColors rgbColor mruColors").Split(' '), StringComparer.Ordinal);

    private static bool KnownFormat(System.Xml.Linq.XElement element, string ns, bool destructiveStyleCleanup)
    {
        var parent = element.Parent?.Name.LocalName;
        if (!KnownStyleElements.Contains(element.Name.LocalName)) return false;
        var inFormat = element.AncestorsAndSelf().Any(e => e.Name.LocalName == "xf" &&
            e.Parent?.Name.LocalName is "cellXfs" or "cellStyleXfs");
        if (!inFormat && parent != "cellStyles") return true;
        var allowed = element.Name.LocalName switch
        {
            "xf" => "numFmtId fontId fillId borderId xfId quotePrefix pivotButton applyNumberFormat applyFont applyFill applyBorder applyAlignment applyProtection",
            "alignment" => "horizontal vertical textRotation wrapText shrinkToFit indent relativeIndent justifyLastLine readingOrder",
            "protection" => "locked hidden",
            "cellStyle" => "name xfId builtinId iLevel hidden customBuiltin",
            _ => ""
        };
        if (allowed.Length == 0 || element.Name.NamespaceName != ns) return false;
        var names = allowed.Split(' ').ToHashSet(StringComparer.Ordinal);
        return element.Attributes().All(a => a.IsNamespaceDeclaration ||
            (a.Name.NamespaceName.Length == 0 && names.Contains(a.Name.LocalName)) ||
            (destructiveStyleCleanup && KnownOfficeMetadata.IsAttribute(element.Name.LocalName, a.Name.NamespaceName, a.Name.LocalName)));
    }

}
