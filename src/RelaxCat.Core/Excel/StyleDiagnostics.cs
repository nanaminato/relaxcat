using System.Text;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

internal static class StyleDiagnostics
{
    internal static IReadOnlyList<BaseStyleReference> References(StyleSheetDocument styles, bool allowDuplicateNamedIdentities = false)
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
        // Office imposes stronger named-style constraints than the base ISO schema.
        // Reject invalid input rather than silently assigning identities or reporting Verified.
        var namedBases = new HashSet<int>();
        var namedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var style in OoxmlXml.Child(styles.Root, "cellStyles")?.Elements() ?? [])
        {
            if (OoxmlXml.IntAttribute(style, "xfId") is not { } baseId || baseId < 0 ||
                baseId >= styles.CellStyleXfs.Count || (!namedBases.Add(baseId) && !allowDuplicateNamedIdentities))
                throw new InvalidDataException("Named styles require valid, distinct base-format references for Office compatibility.");
            if (!allowDuplicateNamedIdentities && (style.Attribute("name") is not { } name || string.IsNullOrEmpty(name.Value) || !namedNames.Add(name.Value)))
                throw new InvalidDataException("Named styles require nonempty, distinct names for Office compatibility.");
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

    internal static IEnumerable<StyleReferenceEdge> Graph(StyleSheetDocument styles)
    {
        foreach (var table in new[] { "cellStyles", "cellStyleXfs", "cellXfs" })
        {
            var elements = OoxmlXml.Child(styles.Root, table)?.Elements().ToArray() ?? [];
            for (var index = 0; index < elements.Length; index++)
            {
                var element = elements[index];
                if (element.Attribute("xfId") is { } reference && int.TryParse(reference.Value, out var baseId))
                    yield return new StyleReferenceEdge(table, index, "cellStyleXfs", baseId, "xfId");
                if (table == "cellStyles") continue;
                foreach (var (attribute, target) in new[] { ("fontId", "fonts"), ("fillId", "fills"),
                    ("borderId", "borders"), ("numFmtId", "numFmts") })
                {
                    var id = OoxmlXml.IntAttribute(element, attribute) ?? 0;
                    yield return new StyleReferenceEdge(table, index,
                        target == "numFmts" && id < 164 ? "builtinNumberFormats" : target, id, attribute);
                }
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
        IReadOnlyList<BaseStyleReference> references, HashSet<int> cellReferences,
        IReadOnlyList<StyleTableChange> changes, bool graphComplete)
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
                used, used is null ? null : Math.Max(0, items.Length - used.Value))
            {
                SemanticDuplicateCount = items.Length - items.Select(item => s.Name.LocalName is "cellXfs" or "cellStyleXfs"
                    ? StyleTableOptimizer.SemanticKey(styles, item) : OoxmlXml.CanonicalKey(item)).Distinct(StringComparer.Ordinal).Count(),
                SafeDuplicateRemovalCount = graphComplete && s.Name.LocalName != "cellStyles" ? changes.Where(c => c.Name == s.Name.LocalName).Sum(c => c.DuplicatesRemoved) : 0,
                ConditionalNamedStyleRemovalCount = s.Name.LocalName == "cellStyles" ? changes.Where(c => c.Name == "cellStyles").Sum(c => c.DuplicatesRemoved) : 0,
                SafeUnusedRemovalCount = graphComplete ? changes.Where(c => c.Name == s.Name.LocalName).Sum(c => c.UnusedRemoved) : 0
            };
        }).ToArray();

        int ComponentReferences(string attribute) => styles.CellXfs.Concat(styles.CellStyleXfs)
            .Select(xf => OoxmlXml.IntAttribute(xf, attribute) ?? 0).Distinct().Count();
    }
}
