using System.Globalization;
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

    private StyleDeduplicationPlan(int[] map, IReadOnlyList<XElement> kept, IReadOnlyList<string> originalKeys,
        XDocument document, byte[] stylesXml)
    {
        Map = map;
        Kept = kept;
        OriginalKeys = originalKeys;
        Document = document;
        StylesXml = stylesXml;
    }

    internal static StyleDeduplicationPlan Build(StyleSheetDocument styles)
    {
        var original = styles.CellXfs;
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

        var clone = new XDocument(styles.Document);
        var cellXfs = OoxmlXml.Child(clone.Root, "cellXfs") ??
                      throw new InvalidDataException("The workbook style sheet has no cell format table (cellXfs).");
        cellXfs.RemoveNodes();
        foreach (var format in kept) cellXfs.Add(format);
        cellXfs.SetAttributeValue("count", kept.Count.ToString(CultureInfo.InvariantCulture));
        return new StyleDeduplicationPlan(map, kept, keys, clone, OoxmlXml.Serialize(clone));
    }
}

internal sealed record WorksheetStyleScan(
    string UsedRange,
    int CellCount,
    Dictionary<int, int> EffectiveStyleCounts,
    Dictionary<int, int> RawStyleCounts,
    List<int> StyleOrder);

internal static class WorksheetStyleScanner
{
    internal static WorksheetStyleScan Scan(XDocument document, int cellXfCount, string sheetName)
    {
        var root = document.Root ?? throw new InvalidDataException($"Worksheet '{sheetName}' has no root element.");
        var columns = ReadColumnStyles(root, cellXfCount, sheetName);
        var effective = new Dictionary<int, int>();
        var raw = new Dictionary<int, int>();
        foreach (var (_, _, style) in columns) Increment(raw, style);
        var order = new List<int>();
        var seen = new HashSet<int>();
        var cellCount = 0;
        var minimumRow = int.MaxValue;
        var maximumRow = 0;
        var minimumColumn = int.MaxValue;
        var maximumColumn = 0;
        var rowIndex = 0;

        foreach (var row in OoxmlXml.Child(root, "sheetData")?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (row.Name.LocalName != "row") continue;
            var rowNumber = OoxmlXml.IntAttribute(row, "r") ?? rowIndex + 1;
            rowIndex = rowNumber;
            var rowStyle = OoxmlXml.IntAttribute(row, "s");
            if (rowStyle is { } declaredRowStyle)
            {
                Validate(declaredRowStyle, cellXfCount, sheetName);
                Increment(raw, declaredRowStyle);
            }

            var columnIndex = 0;
            foreach (var cell in row.Elements())
            {
                if (cell.Name.LocalName != "c") continue;
                var cellRow = rowNumber;
                if (OoxmlXml.Attribute(cell, "r") is { } reference)
                {
                    if (TryParseReference(reference, out var parsedRow, out var parsedColumn))
                    {
                        cellRow = parsedRow;
                        columnIndex = parsedColumn;
                    }
                    else columnIndex++;
                }
                else columnIndex++;

                var ownStyle = OoxmlXml.IntAttribute(cell, "s");
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
        }

        var dimension = OoxmlXml.Attribute(OoxmlXml.Child(root, "dimension"), "ref");
        var usedRange = !string.IsNullOrWhiteSpace(dimension)
            ? dimension
            : cellCount > 0
                ? OoxmlXml.CellReference(minimumRow, minimumColumn) + ":" + OoxmlXml.CellReference(maximumRow, maximumColumn)
                : "";
        return new WorksheetStyleScan(usedRange, cellCount, effective, raw, order);
    }

    private static void Validate(int style, int cellXfCount, string sheetName)
    {
        if (style < 0 || style >= cellXfCount)
            throw new InvalidDataException(
                $"Worksheet '{sheetName}' references cell format {style}, but the workbook defines {cellXfCount} cell formats.");
    }

    private static void Increment(Dictionary<int, int> counts, int style) =>
        counts[style] = counts.GetValueOrDefault(style) + 1;

    private static List<(int Minimum, int Maximum, int Style)> ReadColumnStyles(XElement root, int cellXfCount, string sheetName)
    {
        var result = new List<(int, int, int)>();
        if (OoxmlXml.Child(root, "cols") is not { } columns) return result;
        foreach (var column in OoxmlXml.Children(columns, "col"))
        {
            if (OoxmlXml.IntAttribute(column, "style") is not { } style) continue;
            var minimum = OoxmlXml.IntAttribute(column, "min") ?? 0;
            var maximum = OoxmlXml.IntAttribute(column, "max") ?? 0;
            if (minimum < 1 || maximum < minimum) continue;
            Validate(style, cellXfCount, sheetName);
            result.Add((minimum, maximum, style));
        }
        return result;
    }

    private static int? ColumnStyle(List<(int Minimum, int Maximum, int Style)> columns, int index)
    {
        foreach (var (minimum, maximum, style) in columns)
            if (index >= minimum && index <= maximum) return style;
        return null;
    }

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
}
