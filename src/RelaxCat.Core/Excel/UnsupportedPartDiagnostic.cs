using System.Text.Json.Serialization;

namespace RelaxCat.Core.Excel;

[JsonConverter(typeof(JsonStringEnumConverter<UnsupportedPartKind>))]
public enum UnsupportedPartKind
{
    Other, Drawing, VmlDrawing, Chart, Table, Comment, Pivot, Relationship,
    Media, PrinterSettings, ExternalLink, Control, EmbeddedObject, Signature, CustomXml, Metadata
}

[JsonConverter(typeof(JsonStringEnumConverter<UnsupportedPartReason>))]
public enum UnsupportedPartReason
{
    UnrecognizedPart, MalformedXml, UnexpectedRoot, ExtensionList, AlternateContent,
    UnknownElementNamespace, UnknownAttributeNamespace, UnhandledStyleReference,
    InvalidStyleReference, NamedStyleReference, VmlNonCommentObject, VmlUnsupportedObject,
    RelationshipOwner, RelationshipType, InvalidRelationship
}

/// <summary>Fixed vocabulary only: one first rejection per rejected part, aggregated without names.</summary>
public sealed record UnsupportedPartDiagnostic(UnsupportedPartKind PartKind, UnsupportedPartReason Reason, int PartCount)
{
    internal static UnsupportedPartKind Classify(string part)
    {
        if (part.EndsWith(".rels", StringComparison.Ordinal)) return UnsupportedPartKind.Relationship;
        if (part.StartsWith("xl/drawings/", StringComparison.Ordinal))
            return part.EndsWith(".vml", StringComparison.Ordinal) ? UnsupportedPartKind.VmlDrawing : UnsupportedPartKind.Drawing;
        if (part.StartsWith("xl/charts/", StringComparison.Ordinal)) return UnsupportedPartKind.Chart;
        if (part.StartsWith("xl/tables/", StringComparison.Ordinal)) return UnsupportedPartKind.Table;
        if (part.StartsWith("xl/comments", StringComparison.Ordinal) || part.StartsWith("xl/threadedComments/", StringComparison.Ordinal)) return UnsupportedPartKind.Comment;
        if (part.StartsWith("xl/pivot", StringComparison.Ordinal)) return UnsupportedPartKind.Pivot;
        if (part.StartsWith("xl/media/", StringComparison.Ordinal)) return UnsupportedPartKind.Media;
        if (part.StartsWith("xl/printerSettings/", StringComparison.Ordinal)) return UnsupportedPartKind.PrinterSettings;
        if (part.StartsWith("xl/externalLinks/", StringComparison.Ordinal)) return UnsupportedPartKind.ExternalLink;
        if (part.StartsWith("xl/activeX/", StringComparison.Ordinal) || part.StartsWith("xl/ctrlProps/", StringComparison.Ordinal)) return UnsupportedPartKind.Control;
        if (part.StartsWith("xl/embeddings/", StringComparison.Ordinal)) return UnsupportedPartKind.EmbeddedObject;
        if (part.StartsWith("_xmlsignatures/", StringComparison.Ordinal)) return UnsupportedPartKind.Signature;
        if (part.StartsWith("customXml/", StringComparison.Ordinal)) return UnsupportedPartKind.CustomXml;
        if (part.StartsWith("docProps/", StringComparison.Ordinal) || part == "xl/metadata.xml") return UnsupportedPartKind.Metadata;
        return UnsupportedPartKind.Other;
    }
}
