using System.Xml;

namespace RelaxCat.Core.Excel;

/// <summary>Conservative closed-world gate for table GC/component index changes.</summary>
internal sealed record StyleOptimizationContext(HashSet<int> CellReferences, string? SkipReason)
{
    internal static StyleOptimizationContext Read(OoxmlPackage package, StyleSheetDocument styles,
        IReadOnlyList<WorkbookSheet> worksheets, CancellationToken token, HashSet<int>? knownReferences = null,
        bool allowMacroBehaviorLoss = false)
    {
        var references = knownReferences ?? new HashSet<int>();
        references.Add(0); // Excel's implicit/default cell format is always a root.
        string? reason = StyleDeduplicationPlan.CanMergeBaseStyles(styles) ? null : "styleExtensions";
        foreach (var element in styles.Root.Descendants())
        {
            token.ThrowIfCancellationRequested();
            var parent = element.Parent?.Name.LocalName;
            if (!KnownFormat(element, styles.Root.Name.NamespaceName)) reason ??= "unsupportedReferences";
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
                if (!allowed) reason ??= "unsupportedReferences";
            }
        }
        var parts = new HashSet<string>(worksheets.Select(s => s.PartName), StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
            "xl/styles.xml", "xl/sharedStrings.xml", "xl/calcChain.xml", "xl/theme/theme1.xml",
            "docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"
        };
        // Explicit lossy policy may break macro name/index assumptions. Preserve VBA bytes,
        // but do not allow signed packages, unknown XML consumers, controls or other parts.
        if (allowMacroBehaviorLoss) parts.Add("xl/vbaProject.bin");
        foreach (var entry in package.Archive.Entries)
        {
            token.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith('/')) continue;
            if (!parts.Contains(entry.FullName)) reason ??= "unsupportedParts";
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
            while (reader.Read())
            {
                token.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.NamespaceURI != styles.Root.Name.NamespaceName || reader.LocalName == "extLst")
                    reason ??= "packageExtensions";
                var elementName = reader.LocalName;
                while (reader.MoveToNextAttribute())
                {
                    if (reader.NamespaceURI.Length > 0 && reader.NamespaceURI is not
                        ("http://www.w3.org/2000/xmlns/" or "http://www.w3.org/XML/1998/namespace" or
                         "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
                         "http://purl.oclc.org/ooxml/officeDocument/relationships"))
                        reason ??= "packageExtensions";
                    if (reader.LocalName is "fontId" or "fillId" or "borderId" or "numFmtId" or "xfId" ||
                        (reader.LocalName == "s" && elementName is not ("c" or "row")) ||
                        (reader.LocalName == "style" && elementName != "col"))
                        reason ??= "unsupportedReferences";
                }
                reader.MoveToElement();
            }
        }
        return new StyleOptimizationContext(references, reason);
    }
    private static readonly HashSet<string> KnownStyleElements = new(
        ("numFmts numFmt fonts font name charset family b i strike outline shadow condense extend color sz u vertAlign scheme " +
         "fills fill patternFill fgColor bgColor gradientFill stop borders border left right top bottom diagonal vertical horizontal start end " +
         "cellStyleXfs cellXfs xf alignment protection cellStyles cellStyle dxfs dxf tableStyles tableStyle tableStyleElement " +
         "colors indexedColors rgbColor mruColors").Split(' '), StringComparer.Ordinal);

    private static bool KnownFormat(System.Xml.Linq.XElement element, string ns)
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
            (a.Name.NamespaceName.Length == 0 && names.Contains(a.Name.LocalName)));
    }

}
