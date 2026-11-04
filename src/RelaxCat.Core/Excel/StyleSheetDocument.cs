using System.Globalization;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>
/// Parsed xl/styles.xml plus the workbook theme. Resolves cell formats into previewable
/// descriptions without ever reading cell values or formulas.
/// </summary>
internal sealed class StyleSheetDocument
{
    private static readonly Dictionary<int, string> BuiltinNumberFormats = new()
    {
        [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
        [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
        [14] = "mm-dd-yy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy",
        [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@"
    };

    private static readonly Dictionary<string, double> BorderWidths = new(StringComparer.Ordinal)
    {
        ["hair"] = .5, ["dotted"] = .75, ["dashDotDot"] = .75, ["dashDot"] = .75, ["dashed"] = .75,
        ["thin"] = .9, ["mediumDashDotDot"] = 1.4, ["slantDashDot"] = 1.4, ["mediumDashDot"] = 1.4,
        ["mediumDashed"] = 1.4, ["medium"] = 1.6, ["double"] = 2.2, ["thick"] = 2.2
    };

    // CT_Stylesheet is a sequence, not a choice: Excel repairs files whose sections are out of order.
    private static readonly string[] SectionOrder =
    {
        "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
        "cellStyles", "dxfs", "tableStyles", "colors", "extLst"
    };

    private readonly Dictionary<int, CellStylePreview> previewCache = new();
    private Dictionary<int, string>? namedIdentities;

    // Named style identity is observable through later style edits, even when formatting matches.
    internal string NamedIdentity(int baseId)
    {
        namedIdentities ??= (OoxmlXml.Child(Root, "cellStyles")?.Elements() ?? [])
            .Where(e => OoxmlXml.IntAttribute(e, "xfId") is not null)
            .GroupBy(e => OoxmlXml.IntAttribute(e, "xfId")!.Value)
            .ToDictionary(g => g.Key, g => string.Concat(g.Select(e =>
            {
                var copy = new XElement(e);
                copy.Attribute("xfId")?.Remove();
                var key = OoxmlXml.CanonicalKey(copy);
                return key.Length.ToString(CultureInfo.InvariantCulture) + ":" + key;
            })));
        return namedIdentities.GetValueOrDefault(baseId, "");
    }

    internal byte[] OriginalXml { get; private set; } = Array.Empty<byte>();
    internal XDocument Document { get; }
    internal XElement Root { get; }
    internal List<XElement> Fonts { get; }
    internal List<XElement> Fills { get; }
    internal List<XElement> Borders { get; }
    internal List<XElement> CellStyleXfs { get; }
    internal List<XElement> CellXfs { get; }
    internal Dictionary<int, string> NumberFormats { get; }
    internal ColorResolver Colors { get; }

    private StyleSheetDocument(XDocument document, XElement root, ColorResolver colors)
    {
        Document = document;
        Root = root;
        Colors = colors;
        Fonts = ListOf(root, "fonts", "font");
        Fills = ListOf(root, "fills", "fill");
        Borders = ListOf(root, "borders", "border");
        CellStyleXfs = ListOf(root, "cellStyleXfs", "xf");
        CellXfs = ListOf(root, "cellXfs", "xf");
        NumberFormats = new Dictionary<int, string>();
        if (OoxmlXml.Child(root, "numFmts") is { } numberFormats)
            foreach (var format in OoxmlXml.Children(numberFormats, "numFmt"))
                if (OoxmlXml.IntAttribute(format, "numFmtId") is { } id &&
                    OoxmlXml.Attribute(format, "formatCode") is { } code)
                    NumberFormats[id] = code;
    }

    internal static StyleSheetDocument Load(OoxmlPackage package)
    {
        var document = package.ReadXml("xl/styles.xml");
        var root = document.Root ?? throw new InvalidDataException("The workbook style sheet has no root element.");
        if (root.Name.LocalName != "styleSheet")
            throw new InvalidDataException("The workbook's xl/styles.xml is not a style sheet part.");
        ValidateSectionOrder(root);
        if (OoxmlXml.Child(root, "cellXfs") is null)
            throw new InvalidDataException("The workbook style sheet has no cell format table (cellXfs).");
        var colors = new ColorResolver(
            package.TryReadXml("xl/theme/theme1.xml"),
            OoxmlXml.Child(OoxmlXml.Child(root, "colors"), "indexedColors"));
        using var source = package.OpenPart("xl/styles.xml");
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        var styles = new StyleSheetDocument(document, root, colors) { OriginalXml = buffer.ToArray() };
        if (styles.CellXfs.Count == 0)
            throw new InvalidDataException("The workbook style sheet does not define any cell formats.");
        return styles;
    }

    internal StyleSheetDocument WithXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        return new StyleSheetDocument(document, document.Root!, Colors) { OriginalXml = bytes };
    }

    private static List<XElement> ListOf(XElement root, string containerName, string itemName)
    {
        var container = OoxmlXml.Child(root, containerName);
        return container is null ? new List<XElement>() : OoxmlXml.Children(container, itemName).ToList();
    }

    private static void ValidateSectionOrder(XElement root)
    {
        var previous = -1;
        foreach (var section in root.Elements())
        {
            var current = Array.IndexOf(SectionOrder, section.Name.LocalName);
            if (current < 0) continue;
            if (current < previous)
                throw new InvalidDataException(
                    $"The workbook's xl/styles.xml has '{section.Name.LocalName}' out of the OOXML section order; " +
                    "Excel would repair this workbook, so it is not safe to analyze or optimize.");
            previous = current;
        }
    }

    internal CellStylePreview Preview(int styleId)
    {
        if (styleId < 0 || styleId >= CellXfs.Count)
            throw new InvalidDataException($"Cell format index {styleId} is outside the {CellXfs.Count} formats defined by the workbook.");
        if (previewCache.TryGetValue(styleId, out var cached)) return cached;

        var xf = CellXfs[styleId];
        var baseXf = BaseStyleXf(xf);
        var (fillColor, fillPattern) = ResolveFill(At(Fills, IndexOf(xf, baseXf, "applyFill", "fillId")));
        var (fontName, fontSize, bold, italic, underline, strike, fontColor) =
            ResolveFont(At(Fonts, IndexOf(xf, baseXf, "applyFont", "fontId")));
        var (borders, borderColor, borderThickness) =
            ResolveBorder(At(Borders, IndexOf(xf, baseXf, "applyBorder", "borderId")));
        var numberFormat = NumberFormatCode(IndexOf(xf, baseXf, "applyNumberFormat", "numFmtId"));
        var alignment = OoxmlXml.Attribute(xf, "applyAlignment") is "0" or "false"
            ? OoxmlXml.Child(baseXf, "alignment") ?? OoxmlXml.Child(xf, "alignment")
            : OoxmlXml.Child(xf, "alignment");
        var horizontal = Horizontal(OoxmlXml.Attribute(alignment, "horizontal"));
        var vertical = Vertical(OoxmlXml.Attribute(alignment, "vertical"));
        var wrap = OoxmlXml.Attribute(alignment, "wrapText") is "1" or "true";

        var isDefault = fillColor is null && borders == BorderSides.None && !bold && !italic && !underline && !strike &&
                        string.Equals(fontName, "Calibri", StringComparison.OrdinalIgnoreCase) &&
                        Math.Abs(fontSize - 11) < .01 && numberFormat == "General" &&
                        horizontal == PreviewHorizontalAlignment.General && vertical == PreviewVerticalAlignment.Bottom && !wrap;

        var preview = new CellStylePreview(styleId, isDefault, fillColor, fillPattern, fontName, fontSize,
            bold, italic, underline, strike, fontColor, borders, borderColor, borderThickness,
            horizontal, vertical, wrap, numberFormat);
        previewCache[styleId] = preview;
        return preview;
    }

    private static XElement? At(List<XElement> items, int index) =>
        index >= 0 && index < items.Count ? items[index] : null;

    private XElement? BaseStyleXf(XElement xf)
    {
        var id = OoxmlXml.IntAttribute(xf, "xfId");
        return id is { } value && value >= 0 && value < CellStyleXfs.Count ? CellStyleXfs[value] : null;
    }

    private static int IndexOf(XElement xf, XElement? baseXf, string applyName, string idName)
    {
        var own = OoxmlXml.IntAttribute(xf, idName) ?? 0;
        return OoxmlXml.Attribute(xf, applyName) is "0" or "false" && baseXf is not null
            ? OoxmlXml.IntAttribute(baseXf, idName) ?? own
            : own;
    }

    private string NumberFormatCode(int id) =>
        NumberFormats.TryGetValue(id, out var code) ? code : BuiltinNumberFormats.GetValueOrDefault(id, "General");

    private (string? Color, string Pattern) ResolveFill(XElement? fill)
    {
        if (fill is null) return (null, "none");
        if (OoxmlXml.Child(fill, "patternFill") is { } patternFill)
        {
            var pattern = OoxmlXml.Attribute(patternFill, "patternType") ?? "none";
            if (pattern is "none" or "gray125") return (null, pattern);
            return (Colors.Resolve(OoxmlXml.Child(patternFill, "fgColor")), pattern);
        }
        var gradient = OoxmlXml.Child(fill, "gradientFill");
        if (gradient is null) return (null, "none");
        var stop = gradient.Descendants().FirstOrDefault(element => element.Name.LocalName == "stop");
        return (Colors.Resolve(OoxmlXml.Child(stop, "color")), "gradient");
    }

    private (string Name, double Size, bool Bold, bool Italic, bool Underline, bool Strike, string Color) ResolveFont(XElement? font)
    {
        var name = OoxmlXml.Attribute(OoxmlXml.Child(font, "name"), "val") ?? "Calibri";
        var size = OoxmlXml.DoubleAttribute(OoxmlXml.Child(font, "sz"), "val") ?? 11;
        var color = Colors.Resolve(OoxmlXml.Child(font, "color")) ?? "#000000";
        var underlineElement = OoxmlXml.Child(font, "u");
        var underline = underlineElement is not null &&
                        OoxmlXml.Attribute(underlineElement, "val") is not ("none" or "0" or "false");
        static bool Flag(XElement? element) => element is not null &&
                                               OoxmlXml.Attribute(element, "val") is not ("0" or "false");
        return (name, size, Flag(OoxmlXml.Child(font, "b")), Flag(OoxmlXml.Child(font, "i")),
            underline, Flag(OoxmlXml.Child(font, "strike")), color);
    }

    private (BorderSides Sides, string? Color, double Width) ResolveBorder(XElement? border)
    {
        if (border is null) return (BorderSides.None, null, 0);
        var sides = BorderSides.None;
        string? color = null;
        double width = 0;
        foreach (var (name, side) in new (string Name, BorderSides Side)[]
                 {
                     ("left", BorderSides.Left), ("right", BorderSides.Right),
                     ("top", BorderSides.Top), ("bottom", BorderSides.Bottom)
                 })
        {
            var edge = OoxmlXml.Child(border, name);
            var style = OoxmlXml.Attribute(edge, "style");
            if (style is null or "none") continue;
            sides |= side;
            width = Math.Max(width, BorderWidths.GetValueOrDefault(style, .9));
            color ??= Colors.Resolve(OoxmlXml.Child(edge, "color"));
        }
        if (sides == BorderSides.None) return (sides, null, 0);
        return (sides, color ?? "#000000", width);
    }

    private static PreviewHorizontalAlignment Horizontal(string? value) => value switch
    {
        "left" => PreviewHorizontalAlignment.Left,
        "center" => PreviewHorizontalAlignment.Center,
        "right" => PreviewHorizontalAlignment.Right,
        "fill" => PreviewHorizontalAlignment.Fill,
        "justify" => PreviewHorizontalAlignment.Justify,
        "centerContinuous" => PreviewHorizontalAlignment.CenterContinuous,
        "distributed" => PreviewHorizontalAlignment.Distributed,
        _ => PreviewHorizontalAlignment.General
    };

    private static PreviewVerticalAlignment Vertical(string? value) => value switch
    {
        "top" => PreviewVerticalAlignment.Top,
        "center" => PreviewVerticalAlignment.Center,
        "justify" => PreviewVerticalAlignment.Justify,
        "distributed" => PreviewVerticalAlignment.Distributed,
        _ => PreviewVerticalAlignment.Bottom
    };
}
