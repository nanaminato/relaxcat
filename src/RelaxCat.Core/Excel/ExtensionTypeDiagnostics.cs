using System.Globalization;
using System.Xml;

namespace RelaxCat.Core.Excel;

/// <summary>Local type metadata only. URI and XML names can still contain private custom information.</summary>
public sealed record ExtensionTypeEntry(string Kind, string Name, int Count);
public sealed record ExtensionTypeDiagnostics(IReadOnlyList<ExtensionTypeEntry> Entries, bool Truncated)
{
    public const int MaximumEntries = 64;
    public const int MaximumNameLength = 128;
    public const string PrivacyWarning = "Local extension type metadata only. URI, namespace and XML names may contain private custom information: review before sharing. No text, ordinary attribute values, workbook paths or style names are included. Output is bounded; truncated counts may be partial. Unknown extensions still block optimization.";
}

internal static class ExtensionTypeScanner
{
    internal static ExtensionTypeDiagnostics Read(OoxmlPackage package, IReadOnlyList<WorkbookSheet> sheets,
        string mainNamespace, CancellationToken token)
    {
        var entries = new Dictionary<(string Kind, string Name), int>();
        var truncated = false;
        foreach (var part in new[] { "xl/styles.xml", "xl/workbook.xml" }.Concat(sheets.Select(s => s.PartName)).Distinct())
        {
            using var stream = package.OpenPart(part);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            var extensionDepth = -1;
            while (reader.Read())
            {
                token.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == extensionDepth) extensionDepth = -1;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "extLst" && reader.NamespaceURI == mainNamespace && !reader.IsEmptyElement)
                    extensionDepth = extensionDepth < 0 ? reader.Depth : extensionDepth;
                var extension = extensionDepth >= 0;
                var foreign = reader.NamespaceURI != mainNamespace;
                if (extension || foreign)
                {
                    Add("element", reader.LocalName);
                    Add("namespace", reader.NamespaceURI);
                }
                var isExtension = extension && reader.LocalName == "ext" && reader.NamespaceURI == mainNamespace;
                while (reader.MoveToNextAttribute())
                {
                    if (reader.NamespaceURI == "http://www.w3.org/2000/xmlns/")
                    {
                        if (InterestingNamespace(reader.Value)) Add("namespace", reader.Value);
                        continue;
                    }
                    if (!extension && !foreign && !InterestingNamespace(reader.NamespaceURI)) continue;
                    Add("attribute", reader.LocalName);
                    if (reader.NamespaceURI.Length > 0) Add("namespace", reader.NamespaceURI);
                    // Never inspect other values (including names, paths, formulae and relationship IDs).
                    if (isExtension && reader.LocalName == "uri" && reader.NamespaceURI.Length == 0) Add("uri", reader.Value);
                }
                reader.MoveToElement();
            }
        }
        return new ExtensionTypeDiagnostics(entries.Select(e => new ExtensionTypeEntry(e.Key.Kind, e.Key.Name, e.Value)).ToArray(), truncated);

        bool InterestingNamespace(string ns) => ns.Length > 0 && ns != mainNamespace && ns is not
            ("http://www.w3.org/XML/1998/namespace" or
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships" or
             "http://purl.oclc.org/ooxml/officeDocument/relationships");

        void Add(string kind, string value)
        {
            if (value.Length == 0) return;
            if (value.Length > ExtensionTypeDiagnostics.MaximumNameLength) truncated = true;
            var display = new string(value.Take(ExtensionTypeDiagnostics.MaximumNameLength)
                .Select(c => char.IsControl(c) || char.GetUnicodeCategory(c) is UnicodeCategory.Format or UnicodeCategory.Surrogate or UnicodeCategory.LineSeparator or UnicodeCategory.ParagraphSeparator ? '\uFFFD' : c).ToArray());
            var key = (kind, display);
            if (entries.TryGetValue(key, out var count)) entries[key] = count == int.MaxValue ? count : count + 1;
            else if (entries.Count < ExtensionTypeDiagnostics.MaximumEntries) entries.Add(key, 1);
            else truncated = true;
        }
    }
}
