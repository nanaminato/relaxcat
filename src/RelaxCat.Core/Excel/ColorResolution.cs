using System.Globalization;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

internal readonly record struct RgbColor(byte R, byte G, byte B)
{
    internal string Hex => $"#{R:X2}{G:X2}{B:X2}";

    internal static RgbColor? FromHex(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var hex = value.Trim();
        if (hex.Length == 8) hex = hex[2..]; // ARGB, alpha is not used for previews.
        if (hex.Length != 6) return null;
        return byte.TryParse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r) &&
               byte.TryParse(hex[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g) &&
               byte.TryParse(hex[4..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b)
            ? new RgbColor(r, g, b)
            : null;
    }

    internal RgbColor ApplyTint(double tint)
    {
        if (Math.Abs(tint) < 1e-9) return this;
        ToHls(out var hue, out var luminance, out var saturation);
        luminance = tint < 0 ? luminance * (1 + tint) : luminance * (1 - tint) + tint;
        return FromHls(hue, Math.Clamp(luminance, 0, 1), saturation);
    }

    private void ToHls(out double hue, out double luminance, out double saturation)
    {
        var r = R / 255d;
        var g = G / 255d;
        var b = B / 255d;
        var maximum = Math.Max(r, Math.Max(g, b));
        var minimum = Math.Min(r, Math.Min(g, b));
        luminance = (maximum + minimum) / 2;
        if (maximum - minimum < 1e-9) { hue = 0; saturation = 0; return; }
        var delta = maximum - minimum;
        saturation = luminance > .5 ? delta / (2 - maximum - minimum) : delta / (maximum + minimum);
        if (maximum == r) hue = (g - b) / delta + (g < b ? 6 : 0);
        else if (maximum == g) hue = (b - r) / delta + 2;
        else hue = (r - g) / delta + 4;
        hue /= 6;
    }

    private static RgbColor FromHls(double hue, double luminance, double saturation)
    {
        if (saturation < 1e-9) return FromChannels(luminance, luminance, luminance);
        var q = luminance < .5 ? luminance * (1 + saturation) : luminance + saturation - luminance * saturation;
        var p = 2 * luminance - q;
        return new RgbColor(Channel(p, q, hue + 1d / 3), Channel(p, q, hue), Channel(p, q, hue - 1d / 3));
    }

    private static byte Channel(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        var value = t < 1d / 6 ? p + (q - p) * 6 * t
            : t < .5 ? q
            : t < 2d / 3 ? p + (q - p) * (2d / 3 - t) * 6
            : p;
        return (byte)Math.Round(Math.Clamp(value, 0, 1) * 255);
    }

    private static RgbColor FromChannels(double r, double g, double b) =>
        new((byte)Math.Round(Math.Clamp(r, 0, 1) * 255),
            (byte)Math.Round(Math.Clamp(g, 0, 1) * 255),
            (byte)Math.Round(Math.Clamp(b, 0, 1) * 255));
}

internal sealed class ColorResolver
{
    private static readonly string[] ThemeSchemeNames =
    {
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"
    };

    private static readonly RgbColor?[] ThemeDefaults =
    {
        RgbColor.FromHex("FFFFFF"), RgbColor.FromHex("000000"), RgbColor.FromHex("E7E6E6"), RgbColor.FromHex("44546A"),
        RgbColor.FromHex("4472C4"), RgbColor.FromHex("ED7D31"), RgbColor.FromHex("A5A5A5"), RgbColor.FromHex("FFC000"),
        RgbColor.FromHex("5B9BD5"), RgbColor.FromHex("70AD47"), RgbColor.FromHex("0563C1"), RgbColor.FromHex("954F72")
    };

    // The legacy 56-colour indexed palette plus the two automatic values.
    private static readonly string[] DefaultIndexed =
    {
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
    };

    private readonly RgbColor?[] theme = new RgbColor?[12];
    private readonly RgbColor?[] indexed = new RgbColor?[64];

    internal ColorResolver(XDocument? themeDocument, XElement? indexedColors)
    {
        for (var i = 0; i < ThemeDefaults.Length; i++) theme[i] = ThemeDefaults[i];
        for (var i = 0; i < DefaultIndexed.Length && i < indexed.Length; i++)
            indexed[i] = RgbColor.FromHex(DefaultIndexed[i]);
        LoadTheme(themeDocument);
        LoadIndexed(indexedColors);
    }

    internal string? Resolve(XElement? colorElement)
    {
        if (colorElement is null) return null;
        var tint = Math.Clamp(OoxmlXml.DoubleAttribute(colorElement, "tint") ?? 0, -1, 1);
        if (RgbColor.FromHex(OoxmlXml.Attribute(colorElement, "rgb")) is { } rgb)
            return rgb.ApplyTint(tint).Hex;
        if (OoxmlXml.IntAttribute(colorElement, "indexed") is { } index)
            return index is >= 0 and < 64 && indexed[index] is { } palette ? palette.ApplyTint(tint).Hex : null;
        if (OoxmlXml.IntAttribute(colorElement, "theme") is { } themeIndex)
            return themeIndex is >= 0 and < 12 && theme[themeIndex] is { } themeColor
                ? themeColor.ApplyTint(tint).Hex
                : null;
        return null; // rgb="auto" or an unknown scheme colour.
    }

    private void LoadTheme(XDocument? themeDocument)
    {
        var scheme = themeDocument?.Descendants().FirstOrDefault(element => element.Name.LocalName == "clrScheme");
        if (scheme is null) return;
        for (var i = 0; i < ThemeSchemeNames.Length; i++)
        {
            var name = ThemeSchemeNames[i];
            var entry = scheme.Elements().FirstOrDefault(element => element.Name.LocalName == name);
            var color = entry?.Elements().FirstOrDefault();
            var hex = OoxmlXml.Attribute(color, "val") ?? OoxmlXml.Attribute(color, "lastClr");
            if (RgbColor.FromHex(hex) is { } value) theme[i] = value;
        }
    }

    private void LoadIndexed(XElement? indexedColors)
    {
        if (indexedColors is null) return;
        var position = 0;
        foreach (var entry in indexedColors.Elements().Where(element => element.Name.LocalName == "rgbColor"))
        {
            if (position >= indexed.Length) break;
            if (RgbColor.FromHex(OoxmlXml.Attribute(entry, "rgb")) is { } value) indexed[position] = value;
            position++;
        }
    }
}
