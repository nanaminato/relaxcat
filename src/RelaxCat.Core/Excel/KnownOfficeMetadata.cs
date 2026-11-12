using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>
/// Narrow, shape-checked Office metadata support for explicit destructive cleanup.
/// These records do not consume the remapped style indices. Preserve ancillary
/// part bytes and the structure/values of style metadata;
/// never fetch external links, treat unknown extensions as opaque, or infer support
/// from a familiar namespace or filename alone.
/// </summary>
internal static class KnownOfficeMetadata
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Xr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string Xr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    private const string Xr6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
    private const string Xr10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Calc = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    private const string AbsPath = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
    private const string Slicer = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string Timeline = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    internal static bool CanRemapStyles(StyleSheetDocument styles)
    {
        var root = styles.Root;
        if (root.Name.NamespaceName != Main || styles.CellStyleXfs.Any(x => x.Attribute("xfId") is not null)) return false;
        if (root.Elements().Any(e => e.Name.LocalName is not
            ("numFmts" or "fonts" or "fills" or "borders" or "cellStyleXfs" or "cellXfs" or
             "cellStyles" or "dxfs" or "tableStyles" or "colors" or "extLst"))) return false;
        foreach (var e in root.DescendantsAndSelf())
        {
            if (IsStyleExtension(e)) continue;
            if (e.Name.Namespace != root.Name.Namespace || e.Name.LocalName == "extLst") return false;
            if (e.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.NamespaceName.Length > 0 &&
                !IsAttribute(e.Name.LocalName, a.Name.NamespaceName, a.Name.LocalName))) return false;
        }
        return true;
    }

    internal static bool IsStyleExtension(XElement element)
    {
        var list = element.AncestorsAndSelf().FirstOrDefault(e => e.Name == (XNamespace)Main + "extLst");
        return list?.Parent?.Name == (XNamespace)Main + "styleSheet" && IsStyleExtensionList(list);
    }

    private static bool IsStyleExtensionList(XElement list) => Attributes(list) && NoText(list) &&
        list.Elements().All(ext => Extension(ext, "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
            Slicer, "slicerStyles", "defaultSlicerStyle") ||
            Extension(ext, "{9260A510-F301-46a8-8635-F512D64BE5F5}",
                Timeline, "timelineStyles", "defaultTimelineStyle"));

    private static bool Extension(XElement ext, string uri, string ns, string child, string attribute) =>
        ext.Name == (XNamespace)Main + "ext" && Attributes(ext, "uri") &&
        (string?)ext.Attribute("uri") == uri && NoText(ext) && ext.Elements().Count() == 1 &&
        ext.Elements().Single() is { } body && body.Name == (XNamespace)ns + child &&
        Attributes(body, attribute) && !body.HasElements && NoText(body);

    internal static bool IsAttribute(string element, string ns, string name) =>
        (ns == Mc && name == "Ignorable" && element is "styleSheet" or "workbook" or "worksheet" or "externalLink") ||
        (ns == Xr && name == "uid" && element is "cellStyle" or "worksheet" or "dataValidation" or "table" or "tableColumn" or "autoFilter") ||
        (ns == "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" && name == "uid" && element == "tableColumn") ||
        (ns == Xr2 && name == "uid" && element == "workbookView") ||
        (ns == "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" &&
            ((name == "knownFonts" && element == "fonts") || (name == "dyDescent" && element is "row" or "sheetFormatPr")));

    // Called only for a foreign element/extLst encountered by the streaming scanner.
    // Disposing the subtree reader advances its parent to the matching end element.
    internal static bool ReadPackageElement(XmlReader reader, string part)
    {
        var depth = reader.Depth;
        using var subtree = reader.ReadSubtree();
        var element = XElement.Load(subtree, LoadOptions.PreserveWhitespace);
        if (part != "xl/workbook.xml" || depth != 1) return false;
        if (element.Name == (XNamespace)Mc + "AlternateContent") return AbsolutePathChoice(element, reader);
        if (element.Name == (XNamespace)Xr + "revisionPtr")
            return !element.HasElements && NoText(element) && Attributes(element,
                "revIDLastSave", "documentId", (XNamespace)Xr6 + "coauthVersionLast",
                (XNamespace)Xr6 + "coauthVersionMax", (XNamespace)Xr10 + "uidLastSave");
        if (element.Name != (XNamespace)Main + "extLst" || !Attributes(element) || !NoText(element)) return false;
        return element.Elements().All(ext =>
            ext.Name == (XNamespace)Main + "ext" && Attributes(ext, "uri") &&
            (string?)ext.Attribute("uri") == "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" &&
            NoText(ext) && ext.Elements().Count() == 1 && ext.Elements().Single() is { } body &&
            body.Name == (XNamespace)Calc + "calcFeatures" && Attributes(body) && NoText(body) &&
            body.Elements().All(feature => feature.Name == (XNamespace)Calc + "feature" &&
                Attributes(feature, "name") && feature.Attribute("name") is not null &&
                !feature.HasElements && NoText(feature)));
    }

    private static bool AbsolutePathChoice(XElement element, XmlReader reader)
    {
        if (!Attributes(element) || !NoText(element) || element.Elements().Count() != 1) return false;
        var choice = element.Elements().Single();
        if (choice.Name != (XNamespace)Mc + "Choice" || !Attributes(choice, "Requires") ||
            !NoText(choice) || choice.Elements().Count() != 1) return false;
        var prefix = (string?)choice.Attribute("Requires");
        if (string.IsNullOrEmpty(prefix) || prefix.Any(char.IsWhiteSpace)) return false;
        // Requires can name a prefix declared on workbook, outside this subtree.
        var requiredNamespace = choice.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? reader.LookupNamespace(prefix);
        if (requiredNamespace is not (AbsPath or Timeline)) return false;
        var path = choice.Elements().Single();
        return path.Name == (XNamespace)AbsPath + "absPath" && Attributes(path, "url") &&
               !path.HasElements && NoText(path);
    }

    internal static bool IsAncillaryPart(OoxmlPackage package, string part, IReadOnlyList<WorkbookSheet> sheets)
    {
        if (Matches(part, @"^xl/printerSettings/printerSettings[0-9]+\.bin$")) return true;
        if (Matches(part, @"^xl/externalLinks/externalLink[0-9]+\.xml$")) return ExternalLink(package.ReadXml(part));
        if (Matches(part, @"^xl/externalLinks/_rels/externalLink[0-9]+\.xml\.rels$"))
        {
            var owner = part.Replace("/_rels/", "/", StringComparison.Ordinal)[..^5];
            return package.Find(owner) is not null && Relationships(package.ReadXml(part), external: true, package);
        }
        if (sheets.Any(sheet => part == RelationshipPart(sheet.PartName)))
            return Relationships(package.ReadXml(part), external: false, package);
        return false;
    }

    private static string RelationshipPart(string part)
    {
        var slash = part.LastIndexOf('/');
        return part[..(slash + 1)] + "_rels/" + part[(slash + 1)..] + ".rels";
    }

    private static bool Relationships(XDocument document, bool external, OoxmlPackage package)
    {
        var root = document.Root;
        if (root?.Name != (XNamespace)PackageRel + "Relationships" || !Attributes(root) || !NoText(root)) return false;
        return root.Elements().All(e =>
        {
            if (e.Name != (XNamespace)PackageRel + "Relationship" || e.HasElements || !NoText(e) ||
                !Attributes(e, "Id", "Type", "Target", "TargetMode") ||
                string.IsNullOrEmpty((string?)e.Attribute("Id")) || string.IsNullOrEmpty((string?)e.Attribute("Target"))) return false;
            var target = (string)e.Attribute("Target")!;
            if (external) return (string?)e.Attribute("Type") == Rel + "/externalLinkPath" &&
                                 (string?)e.Attribute("TargetMode") == "External";
            return (string?)e.Attribute("Type") == Rel + "/printerSettings" && e.Attribute("TargetMode") is null &&
                   Matches(target, @"^\.\./printerSettings/printerSettings[0-9]+\.bin$") &&
                   package.Find("xl/" + target[3..]) is not null;
        });
    }

    private static bool ExternalLink(XDocument document)
    {
        var root = document.Root;
        if (root?.Name != (XNamespace)Main + "externalLink") return false;
        foreach (var e in root.DescendantsAndSelf())
        {
            if (e.Name.NamespaceName != Main) return false;
            var parent = e.Parent?.Name.LocalName;
            var allowed = e.Name.LocalName switch
            {
                "externalLink" when e == root => new XName[] { (XNamespace)Mc + "Ignorable" },
                "externalBook" when parent == "externalLink" => [(XNamespace)Rel + "id"],
                "sheetNames" or "definedNames" or "sheetDataSet" when parent == "externalBook" => [],
                "sheetName" when parent == "sheetNames" => ["val"],
                "definedName" when parent == "definedNames" => ["name", "refersTo", "sheetId"],
                "sheetData" when parent == "sheetDataSet" => ["sheetId", "refreshError"],
                "row" when parent == "sheetData" => ["r"],
                "cell" when parent == "row" => ["r", "t"],
                "v" when parent == "cell" => [],
                _ => null
            };
            if (allowed is null || !Attributes(e, allowed)) return false;
            if (e.Name.LocalName != "v" && !NoText(e)) return false;
        }
        return true;
    }

    private static bool Attributes(XElement e, params XName[] allowed) =>
        e.Attributes().All(a => a.IsNamespaceDeclaration || allowed.Contains(a.Name));
    private static bool NoText(XElement e) => e.Nodes().OfType<XText>().All(t => string.IsNullOrWhiteSpace(t.Value));
    private static bool Matches(string value, string expression) => Regex.IsMatch(value, expression, RegexOptions.CultureInvariant);
}
