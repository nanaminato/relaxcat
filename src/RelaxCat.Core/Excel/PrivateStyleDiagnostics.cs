using System.Text.Json;

namespace RelaxCat.Core.Excel;

/// <summary>Sharing-oriented allowlisted projection. Never serialize the source analysis directly.</summary>
public sealed record PrivateStyleDiagnostics(bool LossyMode, bool RemappingBlocked,
    IReadOnlyList<string> BlockerCategories, IReadOnlyList<string> RetainedTableCategories,
    int CellFormats, int BaseFormats, int ProjectedCellFormats, int ProjectedBaseFormats,
    long StylesXmlBytes, long StylesZipBytes, long ProjectedStylesXmlBytes,
    long ProjectedStylesZipBytesEstimate)
{
    public int SchemaVersion => 2;
    public IReadOnlyList<UnsupportedPartDiagnostic> UnsupportedPartDetails { get; init; } = [];
    public string Scope => "All detected gate categories; category presence. UnsupportedPartDetails counts rejected parts by first rejection only, not all defects. No native Excel validation. Projected sizes are estimates.";
    public string Privacy => "Only fixed categories and aggregate counts/sizes; no paths, names, values, formulas, links or raw XML metadata. Aggregate sizes may still be confidential.";

    public static PrivateStyleDiagnostics FromAnalysis(ExcelStyleAnalysis analysis)
    {
        // Do not forward even a supposedly internal reason unchecked: future scanners may attach details.
        var reasons = analysis.AdvancedOptimizationSkipReasons
            .Concat(analysis.AdvancedOptimizationSkipReason is { } first ? [first] : Array.Empty<string>())
            .Select(code => code switch
            {
                "styleExtensions" or "unsupportedReferences" or "unsupportedParts" or "packageExtensions" => code,
                _ => "otherRestriction"
            }).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        var retained = analysis.RetainedStyleTables.Select(code => code switch
        {
            "phoneticFonts" or "differentialNumberFormats" or "ancillaryNumberFormats" => code,
            _ => "otherRetainedTable"
        }).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        return new(analysis.IsLossy, reasons.Length > 0, reasons, retained,
            analysis.CellXfCount, analysis.CellStyleXfCount, analysis.ProjectedCellXfCount,
            analysis.ProjectedCellStyleXfCount, analysis.StylesXmlBytes, analysis.StylesXmlCompressedBytes,
            analysis.ProjectedStylesXmlBytes, analysis.ProjectedStylesXmlCompressedBytesEstimate)
        {
            UnsupportedPartDetails = analysis.UnsupportedPartDetails.Select(d => new UnsupportedPartDiagnostic(
                Enum.IsDefined(d.PartKind) ? d.PartKind : UnsupportedPartKind.Other,
                Enum.IsDefined(d.Reason) ? d.Reason : UnsupportedPartReason.UnrecognizedPart,
                Math.Max(0, d.PartCount))).ToArray()
        };
    }

    public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
}
