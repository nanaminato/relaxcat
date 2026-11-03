using System.Text;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

internal static class StyleDiagnostics
{
    internal static IReadOnlyList<BaseStyleReference> References(StyleSheetDocument styles)
    {
        var numberFormats = OoxmlXml.Child(styles.Root, "numFmts")?.Elements().ToArray() ?? [];
        var ids = new HashSet<int>();
        foreach (var format in numberFormats)
            if (OoxmlXml.IntAttribute(format, "numFmtId") is not { } id || id < 0 || !ids.Add(id) || format.Attribute("formatCode") is null)
                throw new InvalidDataException("Invalid or duplicate custom number format definition.");
        foreach (var section in styles.Root.Elements())
        {
            if (section.Attribute("count") is { } count &&
                (!int.TryParse(count.Value, out var declared) || declared != section.Elements().Count()))
                throw new InvalidDataException($"Invalid {section.Name.LocalName} count '{count.Value}'.");
        }
        foreach (var xf in styles.CellXfs.Concat(styles.CellStyleXfs))
        {
            ValidateComponent(xf, "fontId", styles.Fonts.Count);
            ValidateComponent(xf, "fillId", styles.Fills.Count);
            ValidateComponent(xf, "borderId", styles.Borders.Count);
            if (xf.Attribute("numFmtId") is { } number &&
                (!int.TryParse(number.Value, out var id) || id < 0 || (id >= 164 && !styles.NumberFormats.ContainsKey(id))))
                throw new InvalidDataException($"Invalid numFmtId '{number.Value}'.");
        }
        var cells = new int[styles.CellStyleXfs.Count];
        var names = new int[cells.Length];
        Count(styles.CellXfs, cells);
        Count(OoxmlXml.Child(styles.Root, "cellStyles")?.Elements() ?? [], names);
        return Enumerable.Range(0, cells.Length).Select(i => new BaseStyleReference(i, cells[i], names[i])).ToArray();

        static void Count(IEnumerable<XElement> elements, int[] counts)
        {
            foreach (var element in elements)
            {
                var attribute = element.Attribute("xfId");
                if (attribute is null) continue;
                if (!int.TryParse(attribute.Value, out var id) || id < 0 || id >= counts.Length)
                    throw new InvalidDataException($"Invalid {element.Name.LocalName}.xfId reference '{attribute.Value}'.");
                counts[id]++;
            }
        }
    }

    private static void ValidateComponent(XElement xf, string name, int count)
    {
        if (xf.Attribute(name) is { } attribute &&
            (!int.TryParse(attribute.Value, out var id) || id < 0 || id >= count))
            throw new InvalidDataException($"Invalid {name} '{attribute.Value}'.");
    }

    internal static IReadOnlyList<StyleSectionDiagnostic> Sections(StyleSheetDocument styles,
        IReadOnlyList<BaseStyleReference> references, HashSet<int> cellReferences)
    {
        var sections = styles.Root.Elements().ToArray();
        var sizes = sections.Select(s => (long)Encoding.UTF8.GetByteCount(s.ToString(SaveOptions.DisableFormatting))).ToArray();
        var total = sizes.Sum();
        return sections.Select((s, i) =>
        {
            var items = s.Elements().ToArray();
            int? used = s.Name.LocalName switch
            {
                "cellStyleXfs" => references.Count(r => r.CellXfReferences + r.NamedStyleReferences > 0),
                "cellXfs" => cellReferences.Count,
                "fonts" => ComponentReferences("fontId"),
                "fills" => ComponentReferences("fillId"),
                "borders" => ComponentReferences("borderId"),
                "numFmts" => items.Count(item => styles.CellXfs.Concat(styles.CellStyleXfs)
                    .Any(xf => OoxmlXml.Attribute(xf, "numFmtId") == OoxmlXml.Attribute(item, "numFmtId"))),
                _ => null
            };
            return new StyleSectionDiagnostic(s.Name.LocalName, items.Length, sizes[i],
                total == 0 ? 0 : (double)sizes[i] / total,
                items.Length - items.Select(OoxmlXml.CanonicalKey).Distinct(StringComparer.Ordinal).Count(),
                used, used is null ? null : items.Length - used);
        }).ToArray();

        int ComponentReferences(string attribute) => styles.CellXfs.Concat(styles.CellStyleXfs)
            .Select(xf => OoxmlXml.IntAttribute(xf, attribute) ?? 0).Distinct().Count();
    }
}
