using System.Text.Json;

namespace RelaxCat.Core.Excel;

/// <summary>Sharing-oriented allowlisted projection. Never serialize the source analysis directly.</summary>
public sealed record PrivateStyleDiagnostics(bool LossyMode, bool RemappingBlocked,
    IReadOnlyList<string> BlockerCategories, IReadOnlyList<string> RetainedTableCategories,
    int CellFormats, int BaseFormats, int ProjectedCellFormats, int ProjectedBaseFormats,
    long StylesXmlBytes, long StylesZipBytes, long ProjectedStylesXmlBytes,
    long ProjectedStylesZipBytesEstimate)
{
    public int SchemaVersion => 1;
    public string Scope => "All detected gate categories; category presence, not occurrence counts. No native Excel validation. Projected sizes are estimates.";
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
            analysis.ProjectedStylesXmlBytes, analysis.ProjectedStylesXmlCompressedBytesEstimate);
    }

    public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
}
