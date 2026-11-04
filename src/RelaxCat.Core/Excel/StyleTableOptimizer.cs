using System.Globalization;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>Exact table merges and optional reachability collection. Every edit has an explicit map.</summary>
internal sealed class StyleTableOptimizer
{
    internal StyleSheetDocument Styles { get; private set; }
    internal int[] CellMap { get; private set; }
    internal int[] BaseMap { get; private set; }
    internal int[] NamedMap { get; private set; }
    internal List<StyleTableChange> Changes { get; } = [];
    private readonly CancellationToken token;

    internal StyleTableOptimizer(StyleSheetDocument styles, CancellationToken token)
    {
        Styles = styles;
        CellMap = Enumerable.Range(0, styles.CellXfs.Count).ToArray();
        BaseMap = Enumerable.Range(0, styles.CellStyleXfs.Count).ToArray();
        NamedMap = Enumerable.Range(0, OoxmlXml.Child(styles.Root, "cellStyles")?.Elements().Count() ?? 0).ToArray();
        this.token = token;
    }

    internal void Optimize(ExcelStyleOptimizationOptions options, StyleOptimizationContext? context)
    {
        var advanced = context is { SkipReason: null };
        // Even exact merges renumber IDs. Unknown consumers can therefore make them unsafe.
        if (!advanced) return;
        if (options.RemoveCopiedNamedStyles) RemoveCopiedNames();
        var gc = advanced && options.CollectUnusedStyles;
        if (gc)
        {
            // Preserve cellXfs[0] even for an empty workbook and preserve raw row/column refs.
            RewriteTable("cellXfs", "xf", context!.CellReferences, false, [0]);
            var bases = References(Styles.CellXfs, "xfId");
            bases.UnionWith(References(OoxmlXml.Child(Styles.Root, "cellStyles")?.Elements() ?? [], "xfId"));
            RewriteTable("cellStyleXfs", "xf", bases, false, [0]);
        }
        if (advanced && (options.DeduplicateComponents || gc))
        {
            foreach (var (table, item, attribute, reserved) in new[]
            {
                ("fonts", "font", "fontId", new[] { 0 }),
                ("fills", "fill", "fillId", new[] { 0, 1 }),
                ("borders", "border", "borderId", new[] { 0 })
            })
            {
                var used = gc ? References(Styles.CellXfs.Concat(Styles.CellStyleXfs), attribute) : null;
                RewriteTable(table, item, used, options.DeduplicateComponents, reserved, attribute);
            }
            OptimizeNumberFormats(gc, options.DeduplicateComponents);
        }
        if (StyleDeduplicationPlan.CanMergeBaseStyles(Styles))
            RewriteTable("cellStyleXfs", "xf", null, true, [0]);
        RewriteTable("cellXfs", "xf", null, true, [0]);
    }

    // Opt-in heuristic, never claimed to prove that an external program does not use a name.
    // Require a distinct, existing unsuffixed custom name and identical formatting; any extra
    // metadata, duplicate name, nested suffix or ambiguous name is retained.
    private void RemoveCopiedNames()
    {
        var names = OoxmlXml.Child(Styles.Root, "cellStyles")?.Elements().ToArray() ?? [];
        var byName = names.GroupBy(n => (string?)n.Attribute("name") ?? "", StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.OrdinalIgnoreCase);
        var map = new int[names.Length];
        var kept = 0;
        for (var i = 0; i < names.Length; i++)
        {
            token.ThrowIfCancellationRequested();
            var name = (string?)names[i].Attribute("name") ?? "";
            var match = System.Text.RegularExpressions.Regex.Match(name, @"^(.+)_([1-9][0-9]*)$",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            var remove = false;
            if (Simple(names[i]) && byName[name].Length == 1 && match.Success &&
                !System.Text.RegularExpressions.Regex.IsMatch(match.Groups[1].Value, @"_[1-9][0-9]*$") &&
                byName.TryGetValue(match.Groups[1].Value, out var originals) && originals.Length == 1 && Simple(originals[0]))
            {
                var left = Styles.CellStyleXfs[int.Parse(names[i].Attribute("xfId")!.Value, CultureInfo.InvariantCulture)];
                var right = Styles.CellStyleXfs[int.Parse(originals[0].Attribute("xfId")!.Value, CultureInfo.InvariantCulture)];
                remove = SemanticKey(Styles, left) == SemanticKey(Styles, right);
            }
            map[i] = remove ? -1 : kept++;
        }
        if (kept == names.Length) return;
        NamedMap = map;
        Styles = Styles.WithXml(CellXfsByteRewriter.Rewrite(Styles.OriginalXml, map, "cellStyles", "cellStyle"));
        Changes.Add(new StyleTableChange("cellStyles", names.Length, kept, names.Length - kept, 0));

        static bool Simple(XElement style) => !style.HasElements && style.Attribute("name") is not null &&
            style.Attribute("xfId") is not null && style.Attributes().All(a => a.IsNamespaceDeclaration ||
                (a.Name.NamespaceName.Length == 0 && a.Name.LocalName is "name" or "xfId"));
    }

    private static HashSet<int> References(IEnumerable<XElement> elements, string attribute) =>
        elements.Select(e => OoxmlXml.IntAttribute(e, attribute) ?? 0).ToHashSet();

    private void RewriteTable(string table, string item, HashSet<int>? used, bool deduplicate,
        int[] reserved, string? referenceAttribute = null)
    {
        token.ThrowIfCancellationRequested();
        var section = OoxmlXml.Child(Styles.Root, table);
        if (section is null) return;
        var elements = section.Elements().ToArray();
        var pinned = reserved.Where(id => id < elements.Length).ToHashSet();
        used?.UnionWith(pinned);
        var map = Enumerable.Repeat(-1, elements.Length).ToArray();
        var positions = new Dictionary<string, int>(StringComparer.Ordinal);
        int kept = 0, duplicates = 0, unused = 0;
        for (var i = 0; i < elements.Length; i++)
        {
            token.ThrowIfCancellationRequested();
            if (used is not null && !used.Contains(i)) { unused++; continue; }
            var key = table is "cellXfs" or "cellStyleXfs"
                ? SemanticKey(Styles, elements[i]) : OoxmlXml.CanonicalKey(elements[i]);
            if (deduplicate && !pinned.Contains(i) && positions.TryGetValue(key, out var position))
            { map[i] = position; duplicates++; }
            else { map[i] = kept++; positions.TryAdd(key, map[i]); }
        }
        if (duplicates + unused == 0) return;
        var bytes = CellXfsByteRewriter.Rewrite(Styles.OriginalXml, map, table, item);
        Styles = Styles.WithXml(bytes);
        if (table == "cellXfs")
            CellMap = CellMap.Select(id => id < 0 ? -1 : map[id]).ToArray();
        else
        {
            if (table == "cellStyleXfs") BaseMap = BaseMap.Select(id => id < 0 ? -1 : map[id]).ToArray();
            var attribute = referenceAttribute ?? "xfId";
            var idMap = Enumerable.Range(0, map.Length).ToDictionary(i => i, i => map[i]);
            PatchReferences(attribute, idMap, table == "cellStyleXfs");
        }
        Changes.Add(new StyleTableChange(table, elements.Length, kept, duplicates, unused));
    }

    private void PatchReferences(string attribute, IReadOnlyDictionary<int, int> map, bool baseReferences = false)
    {
        var patches = new Dictionary<(int, string), int>();
        var ordinal = 0;
        foreach (var element in Styles.Root.DescendantsAndSelf())
        {
            var parent = element.Parent?.Name.LocalName;
            var target = element.Name.LocalName == "xf" && parent is "cellXfs" or "cellStyleXfs";
            if (baseReferences) target = (element.Name.LocalName == "xf" && parent == "cellXfs") ||
                                        (element.Name.LocalName == "cellStyle" && parent == "cellStyles");
            if (target && element.Attribute(attribute) is { } reference)
            {
                var id = int.Parse(reference.Value, CultureInfo.InvariantCulture);
                if (map.TryGetValue(id, out var replacement) && replacement != id)
                {
                    if (replacement < 0) throw new InvalidDataException($"Collected a referenced {attribute}.");
                    patches[(ordinal, attribute)] = replacement;
                }
            }
            ordinal++;
        }
        Styles = Styles.WithXml(CellXfsByteRewriter.PatchAttributes(Styles.OriginalXml, patches));
    }

    private void OptimizeNumberFormats(bool gc, bool deduplicate)
    {
        // Differential formats carry independent inline numFmt definitions. Keep custom IDs
        // untouched when these exist; their scope is not the global numFmts index table.
        if (OoxmlXml.Child(Styles.Root, "dxfs")?.Descendants().Any(e => e.Name.LocalName == "numFmt") == true) return;
        var section = OoxmlXml.Child(Styles.Root, "numFmts");
        if (section is null) return;
        var formats = section.Elements().ToArray();
        var ids = formats.Select(f => OoxmlXml.IntAttribute(f, "numFmtId") ?? -1).ToArray();
        if (ids.Any(id => id < 164) || ids.Distinct().Count() != ids.Length) return;
        var used = References(Styles.CellXfs.Concat(Styles.CellStyleXfs), "numFmtId");
        var positions = new Dictionary<string, (int Position, int Id)>(StringComparer.Ordinal);
        var map = Enumerable.Repeat(-1, ids.Length).ToArray();
        var references = new Dictionary<int, int>();
        int kept = 0, duplicates = 0, unused = 0;
        for (var i = 0; i < formats.Length; i++)
        {
            token.ThrowIfCancellationRequested();
            if (gc && !used.Contains(ids[i])) { unused++; continue; }
            var definition = new XElement(formats[i]);
            definition.Attribute("numFmtId")!.Remove();
            var key = OoxmlXml.CanonicalKey(definition);
            if (deduplicate && positions.TryGetValue(key, out var existing))
            { map[i] = existing.Position; references[ids[i]] = existing.Id; duplicates++; }
            else { map[i] = kept++; positions.TryAdd(key, (map[i], ids[i])); }
        }
        if (duplicates + unused == 0) return;
        Styles = Styles.WithXml(CellXfsByteRewriter.Rewrite(Styles.OriginalXml, map, "numFmts", "numFmt"));
        PatchReferences("numFmtId", references);
        Changes.Add(new StyleTableChange("numFmts", formats.Length, kept, duplicates, unused));
    }

    // Only normalize XSD boolean lexical aliases. Never erase absent apply flags, direct
    // formatting, inheritance, protection, or attributes not explicitly understood here.
    private static void NormalizeBooleanLexemes(XElement xf)
    {
        foreach (var element in xf.DescendantsAndSelf())
        {
            if (element.Name.Namespace != xf.Name.Namespace) continue;
            foreach (var attribute in element.Attributes().Where(a => a.Name.NamespaceName.Length == 0))
            {
                var known = element.Name.LocalName switch
                {
                    "xf" => attribute.Name.LocalName is "applyNumberFormat" or "applyFont" or "applyFill" or
                        "applyBorder" or "applyAlignment" or "applyProtection" or "pivotButton" or "quotePrefix",
                    "alignment" => attribute.Name.LocalName is "wrapText" or "shrinkToFit" or "justifyLastLine",
                    "protection" => attribute.Name.LocalName is "locked" or "hidden",
                    _ => false
                };
                if (known && attribute.Value is "true" or "false")
                    attribute.Value = attribute.Value == "true" ? "1" : "0";
            }
        }
    }

    // Resolve IDs to definitions rather than comparing changed indices. Explicit apply flags,
    // inheritance, ordering, and distinct XML remain part of the verification key.
    internal static string SemanticKey(StyleSheetDocument styles, XElement xf)
    {
        XElement Resolve(XElement source, bool includeBase)
        {
            var copy = new XElement(source);
            NormalizeBooleanLexemes(copy);
            foreach (var (name, table) in new[] { ("fontId", styles.Fonts), ("fillId", styles.Fills), ("borderId", styles.Borders) })
                if (copy.Attribute(name) is { } attribute)
                    attribute.Value = OoxmlXml.CanonicalKey(table[int.Parse(attribute.Value, CultureInfo.InvariantCulture)]);
            if (copy.Attribute("numFmtId") is { } number && styles.NumberFormats.TryGetValue(int.Parse(number.Value, CultureInfo.InvariantCulture), out var code))
                number.Value = "custom:" + code;
            if (includeBase && copy.Attribute("xfId") is { } reference)
                reference.Value = OoxmlXml.CanonicalKey(Resolve(styles.CellStyleXfs[int.Parse(reference.Value, CultureInfo.InvariantCulture)], false));
            return copy;
        }
        return OoxmlXml.CanonicalKey(Resolve(xf, true));
    }
}
