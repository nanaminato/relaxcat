using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>Read-only access to an Open Packaging Conventions package (a ZIP of XML parts).</summary>
internal sealed class OoxmlPackage : IDisposable
{
    private readonly FileStream stream;
    private readonly ZipArchive archive;

    internal OoxmlPackage(string path)
    {
        stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        try { archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false); }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    internal ZipArchive Archive => archive;

    internal ZipArchiveEntry? Find(string partName)
    {
        var entry = archive.GetEntry(partName);
        if (entry is not null) return entry;
        return archive.Entries.FirstOrDefault(candidate =>
            string.Equals(candidate.FullName, partName, StringComparison.OrdinalIgnoreCase));
    }

    internal ZipArchiveEntry Require(string partName) =>
        Find(partName) ?? throw new InvalidDataException($"The workbook is missing the required part '{partName}'.");

    internal XDocument ReadXml(string partName)
    {
        using var source = Require(partName).Open();
        using var reader = XmlReader.Create(source, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = true
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    internal Stream OpenPart(string partName) => Require(partName).Open();

    internal XDocument? TryReadXml(string partName)
    {
        var entry = Find(partName);
        if (entry is null) return null;
        using var source = entry.Open();
        using var reader = XmlReader.Create(source, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = true
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    internal long LengthOf(string partName) => Require(partName).Length;

    public void Dispose() => archive.Dispose();
}

internal sealed record WorkbookSheet(string Name, string State, bool IsWorksheet, string PartName);

internal static class WorkbookStructure
{
    internal static IReadOnlyList<WorkbookSheet> ReadSheets(OoxmlPackage package)
    {
        var workbook = package.ReadXml("xl/workbook.xml");
        var relationships = package.ReadXml("xl/_rels/workbook.xml.rels");
        var targets = new Dictionary<string, (string Target, string Type)>(StringComparer.Ordinal);
        foreach (var relationship in relationships.Descendants()
                     .Where(element => element.Name.LocalName == "Relationship"))
        {
            var id = OoxmlXml.Attribute(relationship, "Id");
            var target = OoxmlXml.Attribute(relationship, "Target");
            if (id is null || target is null) continue;
            targets[id] = (target, OoxmlXml.Attribute(relationship, "Type") ?? "");
        }

        var sheets = new List<WorkbookSheet>();
        foreach (var sheet in workbook.Descendants().Where(element => element.Name.LocalName == "sheet"))
        {
            var name = OoxmlXml.Attribute(sheet, "name") ?? "Sheet";
            var state = OoxmlXml.Attribute(sheet, "state") ?? "visible";
            var relationshipId = sheet.Attributes()
                .FirstOrDefault(attribute => attribute.Name.LocalName == "id")?.Value;
            if (relationshipId is null || !targets.TryGetValue(relationshipId, out var relationship)) continue;
            var isWorksheet = relationship.Type.EndsWith("/worksheet", StringComparison.Ordinal);
            sheets.Add(new WorkbookSheet(name, state, isWorksheet, ResolvePartName(relationship.Target)));
        }
        return sheets;
    }

    private static string ResolvePartName(string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var segments = new List<string>();
        foreach (var segment in ("xl/" + target).Split('/'))
        {
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else if (segment is not ("" or ".")) segments.Add(segment);
        }
        return string.Join('/', segments);
    }
}
