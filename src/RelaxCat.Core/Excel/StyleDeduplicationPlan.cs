using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>Builds the old-to-new cell format mapping and the rewritten styles.xml in memory.</summary>
internal sealed class StyleDeduplicationPlan
{
    internal int[] Map { get; }
    internal IReadOnlyList<XElement> Kept { get; }
    internal IReadOnlyList<string> OriginalKeys { get; }
    internal XDocument Document { get; }
    internal byte[] StylesXml { get; }
    internal int KeptCount => Kept.Count;
    internal bool Changed { get; private set; }
    internal static bool CanMergeBaseStyles(StyleSheetDocument styles) =>
        !styles.CellStyleXfs.Any(xf => xf.Attribute("xfId") is not null) &&
        !styles.Root.Elements().Any(e => e.Name.LocalName is not
            ("numFmts" or "fonts" or "fills" or "borders" or "cellStyleXfs" or "cellXfs" or
             "cellStyles" or "dxfs" or "tableStyles" or "colors")) &&
        !styles.Root.Descendants().Any(e => e.Name.Namespace != styles.Root.Name.Namespace || e.Name.LocalName == "extLst") &&
        !styles.Root.DescendantsAndSelf().Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.NamespaceName.Length > 0);


    private StyleDeduplicationPlan(int[] map, IReadOnlyList<XElement> kept, IReadOnlyList<string> originalKeys,
        XDocument document, byte[] stylesXml)
    {
        Map = map;
        Kept = kept;
        OriginalKeys = originalKeys;
        Document = document;
        StylesXml = stylesXml;
    }

    /// <summary>
    /// <paramref name="retainOriginalKeys"/> keeps the canonical key of every input format for
    /// post-write verification. Analysis only needs the projected document, so it can skip the keys.
    /// </summary>
    internal static StyleDeduplicationPlan Build(StyleSheetDocument styles, bool retainOriginalKeys = true)
    {
        // Validate both incoming edges before considering a base-format merge.
        StyleDiagnostics.References(styles);
        var baseMap = Enumerable.Range(0, styles.CellStyleXfs.Count).ToArray();
        if (CanMergeBaseStyles(styles))
        {
            var basePositions = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = 0; i < baseMap.Length; i++)
            {
                var key = OoxmlXml.CanonicalKey(styles.CellStyleXfs[i]);
                if (!basePositions.TryGetValue(key, out var position))
                    basePositions[key] = position = basePositions.Count;
                baseMap[i] = position;
            }
        }
        var baseChanged = baseMap.Distinct().Count() != baseMap.Length;
        var original = styles.CellXfs.Select(xf =>
        {
            var copy = new XElement(xf);
            if (baseChanged && copy.Attribute("xfId") is { } reference)
                reference.Value = baseMap[int.Parse(reference.Value, CultureInfo.InvariantCulture)].ToString(CultureInfo.InvariantCulture);
            return copy;
        }).ToList();
        var keys = original.Select(OoxmlXml.CanonicalKey).ToArray();
        var map = new int[original.Count];
        var positions = new Dictionary<string, int>(StringComparer.Ordinal);
        var kept = new List<XElement>(original.Count);
        for (var i = 0; i < original.Count; i++)
        {
            if (positions.TryGetValue(keys[i], out var existing))
            {
                map[i] = existing;
                continue;
            }
            var position = kept.Count;
            positions[keys[i]] = position;
            map[i] = position;
            kept.Add(new XElement(original[i]));
        }

        // Clone only the sections that stay in place. Cloning the whole document would duplicate
        // tens of thousands of formats just before they are dropped from cellXfs.
        var cloneRoot = new XElement(styles.Root.Name, styles.Root.Attributes());
        foreach (var section in styles.Root.Elements())
        {
            if (section.Name.LocalName == "cellXfs")
            {
                var swapped = new XElement(section.Name,
                    section.Attributes().Where(attribute => attribute.Name.LocalName != "count"));
                swapped.SetAttributeValue("count", kept.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var format in kept) swapped.Add(format);
                cloneRoot.Add(swapped);
            }
            else if (baseChanged && section.Name.LocalName == "cellStyleXfs")
            {
                var copy = new XElement(section);
                var seen = new HashSet<int>();
                var index = 0;
                foreach (var xf in copy.Elements().ToArray())
                    if (!seen.Add(baseMap[index++])) xf.Remove();
                copy.SetAttributeValue("count", seen.Count);
                cloneRoot.Add(copy);
            }
            else
            {
                var copy = new XElement(section);
                if (baseChanged && section.Name.LocalName == "cellStyles")
                    foreach (var style in copy.Elements())
                        if (style.Attribute("xfId") is { } reference)
                            reference.Value = baseMap[int.Parse(reference.Value, CultureInfo.InvariantCulture)].ToString(CultureInfo.InvariantCulture);
                cloneRoot.Add(copy);
            }
        }
        var clone = styles.Document.Declaration is { } declaration
            ? new XDocument(declaration, cloneRoot)
            : new XDocument(cloneRoot);
        var bytes = styles.OriginalXml;
        if (baseChanged)
        {
            bytes = CellXfsByteRewriter.Rewrite(bytes, baseMap, "cellStyleXfs");
            bytes = CellXfsByteRewriter.RemapBaseReferences(bytes, baseMap);
        }
        if (kept.Count != original.Count) bytes = CellXfsByteRewriter.Rewrite(bytes, map);
        return new StyleDeduplicationPlan(map, kept, retainOriginalKeys ? keys : Array.Empty<string>(), clone, bytes)
        { Changed = baseChanged || kept.Count != original.Count };
    }
}

internal sealed record WorksheetStyleScan(
    string UsedRange,
    int CellCount,
    Dictionary<int, int> EffectiveStyleCounts,
    Dictionary<int, int> RawStyleCounts,
    List<int> StyleOrder,
    List<string> MergedRanges);

internal static class WorksheetStyleScanner
{
    /// <summary>
    /// Streaming scan of one worksheet part. Only the style references, used range and counts are
    /// retained, so a sheet with hundreds of thousands of cells never becomes an in-memory tree.
    /// </summary>
    internal static WorksheetStyleScan Scan(Stream stream, int cellXfCount, string sheetName,
        CancellationToken cancellationToken = default, Action<int, int>? rowProgress = null)
    {
        var columns = new List<(int Minimum, int Maximum, int Style)>();
        var effective = new Dictionary<int, int>();
        var raw = new Dictionary<int, int>();
        var order = new List<int>();
        var seen = new HashSet<int>();
        var mergedRanges = new List<string>();
        var cellCount = 0;
        var minimumRow = int.MaxValue;
        var maximumRow = 0;
        var minimumColumn = int.MaxValue;
        var maximumColumn = 0;
        var rowIndex = 0;
        var dimension = (string?)null;
        var expectedRows = 0;
        var rowsSinceProgress = 0;

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = false,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            CheckCharacters = false
        };
        using var reader = XmlReader.Create(stream, settings);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element) continue;
            switch (reader.LocalName)
            {
                case "dimension":
                    dimension = reader.GetAttribute("ref");
                    expectedRows = EndRow(dimension);
                    break;
                case "col":
                    if (ParseInt(reader.GetAttribute("style")) is { } style &&
                        ParseInt(reader.GetAttribute("min")) is { } minimum &&
                        ParseInt(reader.GetAttribute("max")) is { } maximum &&
                        minimum >= 1 && maximum >= minimum)
                    {
                        Validate(style, cellXfCount, sheetName);
                        columns.Add((minimum, maximum, style));
                        Increment(raw, style);
                    }
                    break;
                case "row":
                    ProcessRow();
                    break;
                case "mergeCell":
                    if (reader.GetAttribute("ref") is { Length: > 0 } merged) mergedRanges.Add(merged);
                    break;
            }
        }
        if (rowIndex > 0) rowProgress?.Invoke(rowIndex, expectedRows);

        var usedRange = !string.IsNullOrWhiteSpace(dimension)
            ? dimension
            : cellCount > 0
                ? OoxmlXml.CellReference(minimumRow, minimumColumn) + ":" + OoxmlXml.CellReference(maximumRow, maximumColumn)
                : "";
        return new WorksheetStyleScan(usedRange, cellCount, effective, raw, order, mergedRanges);

        void ProcessRow()
        {
            var rowNumber = ParseInt(reader.GetAttribute("r")) ?? rowIndex + 1;
            rowIndex = rowNumber;
            var rowStyle = ParseInt(reader.GetAttribute("s"));
            if (rowStyle is { } declaredRowStyle)
            {
                Validate(declaredRowStyle, cellXfCount, sheetName);
                Increment(raw, declaredRowStyle);
            }

            var depth = reader.Depth;
            if (reader.IsEmptyElement)
            {
                ReportRow();
                return;
            }

            var columnIndex = 0;
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) break;
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") continue;

                var cellRow = rowNumber;
                if (reader.GetAttribute("r") is { } reference)
                {
                    if (TryParseReference(reference, out var parsedRow, out var parsedColumn))
                    {
                        cellRow = parsedRow;
                        columnIndex = parsedColumn;
                    }
                    else columnIndex++;
                }
                else columnIndex++;

                var ownStyle = ParseInt(reader.GetAttribute("s"));
                if (ownStyle is { } declaredStyle)
                {
                    Validate(declaredStyle, cellXfCount, sheetName);
                    Increment(raw, declaredStyle);
                }
                var style = ownStyle ?? rowStyle ?? ColumnStyle(columns, columnIndex) ?? 0;
                Increment(effective, style);
                if (seen.Add(style)) order.Add(style);
                cellCount++;
                minimumRow = Math.Min(minimumRow, cellRow);
                maximumRow = Math.Max(maximumRow, cellRow);
                minimumColumn = Math.Min(minimumColumn, columnIndex);
                maximumColumn = Math.Max(maximumColumn, columnIndex);
            }
            ReportRow();
        }

        void ReportRow()
        {
            if (rowProgress is null) return;
            if (++rowsSinceProgress < 500) return;
            rowsSinceProgress = 0;
            rowProgress(rowIndex, expectedRows);
        }
    }

    private static void Validate(int style, int cellXfCount, string sheetName)
    {
        if (style < 0 || style >= cellXfCount)
            throw new InvalidDataException(
                $"Worksheet '{sheetName}' references cell format {style}, but the workbook defines {cellXfCount} cell formats.");
    }

    private static void Increment(Dictionary<int, int> counts, int style) =>
        counts[style] = counts.GetValueOrDefault(style) + 1;

    private static int? ColumnStyle(List<(int Minimum, int Maximum, int Style)> columns, int index)
    {
        foreach (var (minimum, maximum, style) in columns)
            if (index >= minimum && index <= maximum) return style;
        return null;
    }

    private static int? ParseInt(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;

    private static bool TryParseReference(string reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        var letters = 0;
        while (letters < reference.Length && char.IsAsciiLetter(reference[letters])) letters++;
        if (letters == 0) return false;
        for (var i = 0; i < letters; i++)
            column = column * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
        return int.TryParse(reference.AsSpan(letters), NumberStyles.None, CultureInfo.InvariantCulture, out row) &&
               row > 0 && column > 0;
    }

    private static int EndRow(string? dimension)
    {
        if (string.IsNullOrWhiteSpace(dimension)) return 0;
        var last = dimension[(dimension.LastIndexOf(':') + 1)..];
        var letters = 0;
        while (letters < last.Length && char.IsAsciiLetter(last[letters])) letters++;
        return int.TryParse(last.AsSpan(letters), NumberStyles.None, CultureInfo.InvariantCulture, out var row) ? row : 0;
    }
}
