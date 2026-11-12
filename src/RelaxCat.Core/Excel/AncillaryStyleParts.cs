using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace RelaxCat.Core.Excel;

/// <summary>Unmodified ordinary OOXML parts, with explicit namespaces and style consumers.
/// Never follow external targets. Unknown extensions/consumers remain a package gate.</summary>
internal static class AncillaryStyleParts
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static bool Accept(OoxmlPackage package, string part, StyleSheetDocument styles,
        List<StyleReferenceEdge> references, CancellationToken token, out UnsupportedPartReason reason)
    {
        reason = UnsupportedPartReason.UnrecognizedPart;
        // An unrecognized/malformed auxiliary part disables edits rather than preventing
        // the diagnostic report or falling through with a partly known reference graph.
        try { return AcceptCore(package, part, styles, references, token); }
        catch (PartRejectedException e) { reason = e.Reason; return false; }
        catch (XmlException) { reason = UnsupportedPartReason.MalformedXml; return false; }
        catch (UriFormatException) { reason = UnsupportedPartReason.InvalidRelationship; return false; }
    }

    private sealed class PartRejectedException(UnsupportedPartReason reason) : Exception
    {
        internal UnsupportedPartReason Reason { get; } = reason;
    }

    private static bool Reject(UnsupportedPartReason reason) => throw new PartRejectedException(reason);

    private static bool AcceptCore(OoxmlPackage package, string part, StyleSheetDocument styles,
        List<StyleReferenceEdge> references, CancellationToken token)
    {
        bool Match(string pattern) => Regex.IsMatch(part, pattern, RegexOptions.CultureInvariant);
        if (Match(@"^xl/media/[^/]+\.(png|jpe?g|gif|bmp|tiff?|emf|wmf)$")) return true;
        if (Match(@"^xl/drawings/[^/]+\.vml$")) return VmlComments(package.ReadXml(part));
        if (part.EndsWith(".rels", StringComparison.Ordinal))
            return Relationships(package, part);
        string? ns = Match(@"^xl/tables/table[0-9]+\.xml$") || Match(@"^xl/(comments[0-9]+|comments/comment[0-9]+)\.xml$") ||
            Match(@"^xl/pivotTables/pivotTable[0-9]+\.xml$") ||
            Match(@"^xl/pivotCache/pivotCache(Definition|Records)[0-9]+\.xml$") ? Main :
            Match(@"^xl/drawings/drawing[0-9]+\.xml$") ? SheetDrawing :
            Match(@"^xl/charts/chart[0-9]+\.xml$") ? Chart : null;
        if (ns is null) return Reject(UnsupportedPartReason.UnrecognizedPart);
        var root = package.ReadXml(part).Root;
        if (root is null || root.Name.NamespaceName != ns || root.Name.LocalName !=
            (ns == SheetDrawing ? "wsDr" : ns == Chart ? "chartSpace" : part.StartsWith("xl/tables/", StringComparison.Ordinal) ? "table" :
             part.StartsWith("xl/pivotTables/", StringComparison.Ordinal) ? "pivotTableDefinition" :
             part.Contains("pivotCacheDefinition", StringComparison.Ordinal) ? "pivotCacheDefinition" :
             part.Contains("pivotCacheRecords", StringComparison.Ordinal) ? "pivotCacheRecords" : "comments")) return Reject(UnsupportedPartReason.UnexpectedRoot);
        // Only a shape's precise creation-identity metadata is known not to consume
        // cell style indices. Preserve bytes; never trust an extension namespace alone.
        var knownMetadata = root.Descendants().Where(IsDrawingCreationIdentity)
            .SelectMany(e => e.DescendantsAndSelf()).ToHashSet();
        var ordinal = -1;
        foreach (var element in root.DescendantsAndSelf())
        {
            token.ThrowIfCancellationRequested();
            ordinal++;
            if (knownMetadata.Contains(element)) continue;
            var elementNs = element.Name.NamespaceName;
            if (element.Name.LocalName == "extLst") return Reject(UnsupportedPartReason.ExtensionList);
            if (element.Name.LocalName == "AlternateContent") return Reject(UnsupportedPartReason.AlternateContent);
            if (!(elementNs == ns || (ns is SheetDrawing or Chart && elementNs is Drawing or Chart)))
                return Reject(UnsupportedPartReason.UnknownElementNamespace);
            foreach (var a in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            {
                if (a.Name.NamespaceName.Length > 0 && a.Name.NamespaceName != Rel &&
                    a.Name.NamespaceName != XNamespace.Xml.NamespaceName &&
                    !KnownOfficeMetadata.IsAttribute(element.Name.LocalName, a.Name.NamespaceName, a.Name.LocalName)) return Reject(UnsupportedPartReason.UnknownAttributeNamespace);
                // DrawingML uses independent typefaces/styles, not workbook font indices.
                // Only the explicit global consumers below are admitted.
                if (a.Name.LocalName is "xfId" or "fontId" or "fillId" or "borderId" ||
                    a.Name.LocalName == "s" || a.Name.LocalName == "style") return Reject(UnsupportedPartReason.UnhandledStyleReference);
                if (a.Name.LocalName.EndsWith("DxfId", StringComparison.Ordinal) || a.Name.LocalName == "dxfId")
                {
                    if (ns != Main || !int.TryParse(a.Value, out var id) || id < 0 ||
                        id >= (OoxmlXml.Child(styles.Root, "dxfs")?.Elements().Count() ?? 0)) return Reject(UnsupportedPartReason.InvalidStyleReference);
                    references.Add(new StyleReferenceEdge(part, ordinal, "dxfs", id, a.Name.LocalName));
                }
                if (a.Name.LocalName is "numFmtId" or "dataCellStyle" or "headerRowCellStyle" or "totalsRowCellStyle" or
                    "pageStyle" or "pivotTableStyle" or "vacatedStyle")
                {
                    // Named style associations cannot survive destructive collapse.
                    if (a.Name.LocalName != "numFmtId") return Reject(UnsupportedPartReason.NamedStyleReference);
                    if (!int.TryParse(a.Value, out var id) || id < 0 || (id >= 164 && !styles.NumberFormats.ContainsKey(id))) return Reject(UnsupportedPartReason.InvalidStyleReference);
                    references.Add(new StyleReferenceEdge(part, ordinal, "numFmts", id, a.Name.LocalName));
                }
            }
        }
        return true;
    }

    private static bool IsDrawingCreationIdentity(XElement list)
    {
        XNamespace a = Drawing;
        XNamespace a16 = "http://schemas.microsoft.com/office/drawing/2014/main";
        if (list.Name != a + "extLst" || list.Parent?.Name != (XNamespace)SheetDrawing + "cNvPr" ||
            list.Attributes().Any(x => !x.IsNamespaceDeclaration) || list.Elements().Count() != 1) return false;
        var ext = list.Elements().Single();
        if (ext.Name != a + "ext" || (string?)ext.Attribute("uri") != "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" ||
            ext.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "uri") || ext.Elements().Count() != 1) return false;
        var id = ext.Elements().Single();
        return id.Name == a16 + "creationId" && !id.HasElements &&
            id.Attributes().Count(x => !x.IsNamespaceDeclaration) == 1 &&
            Guid.TryParseExact((string?)id.Attribute("id"), "B", out _) &&
            list.DescendantNodes().OfType<XText>().All(t => string.IsNullOrWhiteSpace(t.Value));
    }

    // Legacy comment geometry uses VML/CSS styling, not workbook cell styles.
    // Reject controls/OLE and extensions; only Note client data is supported.
    private static bool VmlComments(XDocument document)
    {
        var root = document.Root;
        if (root?.Name != "xml") return Reject(UnsupportedPartReason.UnexpectedRoot);
        var namespaces = new[] { "urn:schemas-microsoft-com:vml", "urn:schemas-microsoft-com:office:office", "urn:schemas-microsoft-com:office:excel" };
        foreach (var element in root.Descendants())
        {
            var textBoxHtml = element.Ancestors().Any(e => e.Name == XName.Get("textbox", namespaces[0])) &&
                element.Name.NamespaceName is "" or "http://www.w3.org/1999/xhtml" &&
                element.Name.LocalName is "div" or "span" or "p" or "br" or "b" or "i" or "u" or "font";
            if (!namespaces.Contains(element.Name.NamespaceName) && !textBoxHtml) return Reject(UnsupportedPartReason.UnknownElementNamespace);
            if (element.Name.LocalName == "ClientData" && (string?)element.Attribute("ObjectType") != "Note") return Reject(UnsupportedPartReason.VmlNonCommentObject);
            if (element.Name.LocalName is "OLEObject" or "control" or "AlternateContent" or "extLst") return Reject(UnsupportedPartReason.VmlUnsupportedObject);
            foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            {
                if (attribute.Name.NamespaceName.Length > 0 && !namespaces.Contains(attribute.Name.NamespaceName) && attribute.Name.NamespaceName != Rel) return Reject(UnsupportedPartReason.UnknownAttributeNamespace);
                if (attribute.Name.LocalName is "fontId" or "fillId" or "borderId" or "xfId" or "numFmtId") return Reject(UnsupportedPartReason.UnhandledStyleReference);
            }
        }
        return true;
    }

    private static bool Relationships(OoxmlPackage package, string part)
    {
        var marker = part.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (marker < 0) return Reject(UnsupportedPartReason.RelationshipOwner);
        var owner = part[..marker] + "/" + part[(marker + 7)..^5];
        if (package.Find(owner) is null) return Reject(UnsupportedPartReason.RelationshipOwner);
        // Only relationships owned by supported worksheets/drawings/charts/tables.
        if (!Regex.IsMatch(owner, @"^xl/(worksheets/[^/]+\.xml|drawings/(drawing[0-9]+\.xml|[^/]+\.vml)|charts/chart[0-9]+\.xml|tables/table[0-9]+\.xml|pivotTables/pivotTable[0-9]+\.xml|pivotCache/pivotCache(Definition|Records)[0-9]+\.xml)$")) return Reject(UnsupportedPartReason.RelationshipOwner);
        var root = package.ReadXml(part).Root;
        if (root?.Name != (XNamespace)PackageRel + "Relationships") return Reject(UnsupportedPartReason.InvalidRelationship);
        return root.Elements().All(e =>
        {
            if (e.Name != (XNamespace)PackageRel + "Relationship" || e.HasElements ||
                e.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.NamespaceName.Length > 0 ||
                    a.Name.LocalName is not ("Id" or "Type" or "Target" or "TargetMode")))) return Reject(UnsupportedPartReason.InvalidRelationship);
            var type = (string?)e.Attribute("Type");
            var target = (string?)e.Attribute("Target");
            if (string.IsNullOrEmpty((string?)e.Attribute("Id")) || string.IsNullOrEmpty(target)) return Reject(UnsupportedPartReason.InvalidRelationship);
            if ((string?)e.Attribute("TargetMode") == "External") return type == Rel + "/hyperlink" || type == Rel + "/image" || Reject(UnsupportedPartReason.RelationshipType);
            if (e.Attribute("TargetMode") is not null || type is null || !type.StartsWith(Rel + "/", StringComparison.Ordinal)) return Reject(UnsupportedPartReason.InvalidRelationship);
            if (type[(Rel.Length + 1)..] is not ("drawing" or "image" or "chart" or "table" or "pivotTable" or "pivotCacheDefinition" or "pivotCacheRecords" or "comments" or "vmlDrawing" or "printerSettings")) return Reject(UnsupportedPartReason.RelationshipType);
            if (target.Contains('\\') || target.Contains('?') || target.Contains('#')) return Reject(UnsupportedPartReason.InvalidRelationship);
            var uri = new Uri(new Uri("https://package.invalid/" + owner), target);
            return (uri.Host == "package.invalid" && package.Find(Uri.UnescapeDataString(uri.AbsolutePath[1..])) is not null) || Reject(UnsupportedPartReason.InvalidRelationship);
        });
    }
}
