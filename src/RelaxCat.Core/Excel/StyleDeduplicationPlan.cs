using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>Builds the old-to-new cell format mapping and the rewritten styles.xml in memory.</summary>
internal sealed class StyleDeduplicationPlan
{
    internal int[] Map { get; }
    internal int[] BaseMap { get; private set; } = [];
    internal int[] NamedMap { get; private set; } = [];
    internal IReadOnlyList<XElement> Kept { get; }
    internal IReadOnlyList<string> OriginalKeys { get; }
    internal XDocument Document { get; }
    internal byte[] StylesXml { get; }
    internal int KeptCount => Kept.Count;
    internal bool Changed { get; private set; }
    internal IReadOnlyList<string> OutputKeys { get; private set; } = [];
    internal IReadOnlyList<StyleTableChange> Changes { get; private set; } = [];
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
    internal static StyleDeduplicationPlan Build(StyleSheetDocument styles, bool retainOriginalKeys = true,
        StyleOptimizationContext? context = null, ExcelStyleOptimizationOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        StyleDiagnostics.References(styles);
        var inputKeys = retainOriginalKeys
            ? styles.CellXfs.Select(xf => StyleTableOptimizer.SemanticKey(styles, xf)).ToArray()
            : Array.Empty<string>();
        var optimizer = new StyleTableOptimizer(styles, cancellationToken);
        optimizer.Optimize(options ?? new ExcelStyleOptimizationOptions(), context);
        var output = optimizer.Styles;
        StyleDiagnostics.References(output);
        var outputKeys = retainOriginalKeys ? output.CellXfs.Select(xf => StyleTableOptimizer.SemanticKey(output, xf)).ToArray() : [];
        if (retainOriginalKeys)
        {
            for (var i = 0; i < optimizer.CellMap.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var mapped = optimizer.CellMap[i];
                if (mapped >= 0 && inputKeys[i] != outputKeys[mapped])
                    throw new InvalidDataException("A retained cell format changed meaning during table optimization.");
            }
            var oldNames = OoxmlXml.Child(styles.Root, "cellStyles")?.Elements().ToArray() ?? [];
            var newNames = OoxmlXml.Child(output.Root, "cellStyles")?.Elements().ToArray() ?? [];
            if (optimizer.NamedMap.Count(id => id >= 0) != newNames.Length)
                throw new InvalidDataException("Named-style mapping is incomplete.");
            for (var i = 0; i < oldNames.Length; i++)
                if (optimizer.NamedMap[i] >= 0 && NamedKey(styles, oldNames[i]) != NamedKey(output, newNames[optimizer.NamedMap[i]]))
                    throw new InvalidDataException("A named style changed meaning during table optimization.");
        }
        return new StyleDeduplicationPlan(optimizer.CellMap, output.CellXfs, inputKeys, output.Document, output.OriginalXml)
        {
            Changed = optimizer.Changes.Count > 0,
            BaseMap = optimizer.BaseMap,
            NamedMap = optimizer.NamedMap,
            Changes = optimizer.Changes.GroupBy(change => change.Name).Select(group => new StyleTableChange(
                group.Key, group.First().Before, group.Last().After, group.Sum(c => c.DuplicatesRemoved), group.Sum(c => c.UnusedRemoved))).ToArray(),
            OutputKeys = outputKeys
        };
    }

    private static string NamedKey(StyleSheetDocument styles, XElement style)
    {
        var copy = new XElement(style);
        if (copy.Attribute("xfId") is { } attribute)
        {
            var reference = new XElement(styles.Root.Name.Namespace + "xf", new XAttribute("xfId", attribute.Value));
            attribute.Value = StyleTableOptimizer.SemanticKey(styles, reference);
        }
        return OoxmlXml.CanonicalKey(copy);
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
                    if (ParseStyle(reader.GetAttribute("style")) is { } style &&
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

        int? ParseStyle(string? value)
        {
            if (value is null) return null;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
                throw new InvalidDataException($"Worksheet '{sheetName}' has an invalid explicit style reference '{value}'.");
            Validate(id, cellXfCount, sheetName);
            return id;
        }

        void ProcessRow()
        {
            var rowNumber = ParseInt(reader.GetAttribute("r")) ?? rowIndex + 1;
            rowIndex = rowNumber;
            var rowStyle = ParseStyle(reader.GetAttribute("s"));
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

                var ownStyle = ParseStyle(reader.GetAttribute("s"));
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
