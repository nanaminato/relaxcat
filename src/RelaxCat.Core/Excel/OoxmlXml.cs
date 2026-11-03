using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

internal static class OoxmlXml
{
    internal static string? Attribute(XElement? element, string localName) =>
        element?.Attributes().FirstOrDefault(attribute =>
            attribute.Name.LocalName == localName && !attribute.IsNamespaceDeclaration)?.Value;

    internal static int? IntAttribute(XElement? element, string localName) =>
        int.TryParse(Attribute(element, localName), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : null;

    internal static double? DoubleAttribute(XElement? element, string localName) =>
        double.TryParse(Attribute(element, localName), NumberStyles.Float, CultureInfo.InvariantCulture, out var value) &&
        double.IsFinite(value) ? value : null;

    internal static IEnumerable<XElement> Children(XElement element, string localName) =>
        element.Elements().Where(child => child.Name.LocalName == localName);

    internal static XElement? Child(XElement? element, string localName) =>
        element?.Elements().FirstOrDefault(child => child.Name.LocalName == localName);

    /// <summary>
    /// Structural key for a cell format. Attribute order and insignificant whitespace are ignored,
    /// so two formats with identical XML always compare equal and two different formats never do.
    /// </summary>
    internal static string CanonicalKey(XElement element)
    {
        var builder = new StringBuilder(128);
        Append(element, builder);
        return builder.ToString();

        static void Append(XElement node, StringBuilder builder)
        {
            static void Field(StringBuilder result, string value) => result.Append(value.Length).Append(':').Append(value);
            builder.Append('<');
            Field(builder, node.Name.NamespaceName);
            Field(builder, node.Name.LocalName);
            foreach (var attribute in node.Attributes()
                         .Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                         .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal))
            {
                Field(builder, attribute.Name.NamespaceName);
                Field(builder, attribute.Name.LocalName);
                Field(builder, attribute.Value);
            }
            builder.Append('[');
            foreach (var child in node.Nodes())
                if (child is XElement element) Append(element, builder);
                else if (child is XText text && !string.IsNullOrWhiteSpace(text.Value))
                { builder.Append('T'); Field(builder, text.Value); }
            builder.Append(']');
        }
    }

    internal static byte[] Serialize(XDocument document)
    {
        using var stream = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            CloseOutput = false
        };
        using (var writer = XmlWriter.Create(stream, settings)) document.Save(writer);
        return stream.ToArray();
    }

    internal static string CellReference(int row, int column)
    {
        var letters = new StringBuilder(3);
        for (var value = column; value > 0; value = (value - 1) / 26) letters.Insert(0, (char)('A' + (value - 1) % 26));
        return letters.Append(row.ToString(CultureInfo.InvariantCulture)).ToString();
    }
}
