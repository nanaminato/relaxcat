using System.Xml;

namespace RelaxCat.Core.Excel;

/// <summary>Conservative closed-world gate for table GC/component index changes.</summary>
internal sealed record StyleOptimizationContext(HashSet<int> CellReferences, string? SkipReason)
{
    internal static StyleOptimizationContext Read(OoxmlPackage package, StyleSheetDocument styles,
        IReadOnlyList<WorkbookSheet> worksheets, CancellationToken token, HashSet<int>? knownReferences = null)
    {
        var references = knownReferences ?? new HashSet<int>();
        references.Add(0); // Excel's implicit/default cell format is always a root.
        string? reason = StyleDeduplicationPlan.CanMergeBaseStyles(styles) ? null : "styleExtensions";
        foreach (var element in styles.Root.Descendants())
        {
            token.ThrowIfCancellationRequested();
            var parent = element.Parent?.Name.LocalName;
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
}
