using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>Read-only analysis of an Excel OOXML workbook's style table and worksheet style usage.</summary>
public sealed class ExcelStyleAnalyzer
{
    /// <summary>Bounds the number of style swatches materialized per worksheet for the UI.</summary>
    internal const int MaximumStylesPerWorksheet = 400;

    public static bool IsSupported(string path)
    {
        var extension = Path.GetExtension(path);
        return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    public ExcelStyleAnalysis Analyze(string path, CancellationToken cancellationToken = default,
        IProgress<ExcelAnalysisProgress>? progress = null, ExcelStyleOptimizationOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A workbook path is required.", nameof(path));
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The workbook was not found.", fullPath);
        if (!IsSupported(fullPath))
            throw new NotSupportedException("Only .xlsx and .xlsm workbooks are supported; .xls is a different binary format.");

        using var package = new OoxmlPackage(fullPath);
        progress?.Report(new ExcelAnalysisProgress(ExcelAnalysisPhase.ReadingStyles, 0, 1));
        var styles = StyleSheetDocument.Load(package);
        var sheetDefinitions = WorkbookStructure.ReadSheets(package).Where(sheet => sheet.IsWorksheet).ToList();
        cancellationToken.ThrowIfCancellationRequested();

        var referenced = new HashSet<int>();
        var summaries = new List<WorksheetStyleSummary>();
        var graph = StyleDiagnostics.Graph(styles).ToList();
        graph.Add(new StyleReferenceEdge("workbook", null, "cellXfs", 0, "implicitDefault"));

        for (var index = 0; index < sheetDefinitions.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = sheetDefinitions[index];
            progress?.Report(new ExcelAnalysisProgress(ExcelAnalysisPhase.ScanningWorksheets,
                index, sheetDefinitions.Count, sheet.Name));
            WorksheetStyleScan scan;
            using (var stream = package.OpenPart(sheet.PartName))
            {
                scan = WorksheetStyleScanner.Scan(stream, styles.CellXfs.Count, sheet.Name, cancellationToken,
                    (rows, expectedRows) => progress?.Report(new ExcelAnalysisProgress(
                        ExcelAnalysisPhase.ScanningWorksheets, index, sheetDefinitions.Count, sheet.Name,
                        rows, expectedRows)));
            }
            foreach (var style in scan.EffectiveStyleCounts.Keys) referenced.Add(style);
            foreach (var style in scan.RawStyleCounts.Keys) referenced.Add(style);
            foreach (var (style, count) in scan.RawStyleCounts)
                graph.Add(new StyleReferenceEdge(sheet.PartName, null, "cellXfs", style, "explicit", count));
            foreach (var (style, count) in scan.EffectiveStyleCounts)
                graph.Add(new StyleReferenceEdge(sheet.PartName, null, "cellXfs", style, "effectiveCells", count));

            var listed = scan.StyleOrder.Take(MaximumStylesPerWorksheet).ToArray();
            var usages = new List<StyleUsage>(listed.Length);
            foreach (var style in listed)
                usages.Add(new StyleUsage(style, scan.EffectiveStyleCounts.GetValueOrDefault(style), styles.Preview(style)));
            summaries.Add(new WorksheetStyleSummary(sheet.Name, sheet.State != "visible", scan.UsedRange,
                scan.CellCount, scan.StyleOrder.Count, Math.Max(0, scan.StyleOrder.Count - listed.Length), usages));
            progress?.Report(new ExcelAnalysisProgress(ExcelAnalysisPhase.ScanningWorksheets,
                index + 1, sheetDefinitions.Count, sheet.Name));
        }

        progress?.Report(new ExcelAnalysisProgress(ExcelAnalysisPhase.Projecting, 0, 1));
        var destructive = options?.DestructiveStyleCleanup == true;
        var context = StyleOptimizationContext.Read(package, styles, sheetDefinitions, cancellationToken, new HashSet<int>(referenced),
            allowMacroBehaviorLoss: destructive);
        var plan = StyleDeduplicationPlan.Build(styles, retainOriginalKeys: false, context: context, options: options,
            cancellationToken: cancellationToken);
        // Equivalence diagnostics are independent of whether unknown consumers prevent edits.
        var uniqueCount = styles.CellXfs.Select(xf => StyleTableOptimizer.SemanticKey(styles, xf))
            .Distinct(StringComparer.Ordinal).Count();
        cancellationToken.ThrowIfCancellationRequested();
        var originalStylesBytes = package.LengthOf("xl/styles.xml");
        // With nothing to merge, optimization copies styles.xml verbatim instead of re-serializing it.
        var projectedStylesBytes = plan.StylesXml.Length;
        return new ExcelStyleAnalysis(fullPath, new FileInfo(fullPath).Length,
            IsMacroEnabled(package, fullPath),
            styles.Fonts.Count, styles.Fills.Count, styles.Borders.Count, styles.CellStyleXfs.Count,
            styles.CellXfs.Count, uniqueCount, referenced.Count,
            originalStylesBytes, projectedStylesBytes, summaries)
        {
            IsLossy = destructive,
            VerificationScope = destructive ? ExcelStyleOptimizationOptions.DestructiveVerificationScope : "SemanticEquivalenceAndPackageIntegrity",
            Warnings = destructive ? [ExcelStyleOptimizationOptions.DestructiveStyleRisk] : [],
            ReferenceGraph = graph,
            ExtensionDiagnostics = ExtensionTypeScanner.Read(package, sheetDefinitions, styles.Root.Name.NamespaceName, cancellationToken),
            NamedStyleCleanupCandidates = (OoxmlXml.Child(styles.Root, "cellStyles")?.Elements() ?? [])
                .Where((_, i) => plan.NamedMap[i] < 0).Select(e => (string?)e.Attribute("name") ?? "").ToArray(),
            ProjectedStylesXmlCompressedBytesEstimate = CompressedSize(plan.StylesXml),
            ProjectedCellXfCount = plan.KeptCount,
            ProjectedCellStyleXfCount = OoxmlXml.Child(plan.Document.Root, "cellStyleXfs")?.Elements().Count() ?? 0,
            TableChanges = plan.Changes,
            AdvancedOptimizationSkipReason = context.SkipReason,
            StylesXmlCompressedBytes = package.Require("xl/styles.xml").CompressedLength,
            BaseStyleReferences = StyleDiagnostics.References(styles, destructive),
            Sections = StyleDiagnostics.Sections(styles, StyleDiagnostics.References(styles, destructive), context.CellReferences,
                plan.Changes, context.SkipReason is null && !destructive),
            UniqueCellStyleXfCount = styles.CellStyleXfs.Select(xf => StyleTableOptimizer.SemanticKey(styles, xf)).Distinct(StringComparer.Ordinal).Count(),
            CanDeduplicateBaseStyles = context.SkipReason is null
        };
    }

    private static long CompressedSize(byte[] bytes)
    {
        using var output = new MemoryStream();
        using (var deflate = new System.IO.Compression.DeflateStream(output,
                   System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(bytes);
        return output.Length;
    }

    internal static bool IsMacroEnabled(OoxmlPackage package, string path) =>
        string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase) ||
        package.Find("xl/vbaProject.bin") is not null;

    internal static IReadOnlyList<WorkbookSheet> ReadWorksheets(OoxmlPackage package) =>
        WorkbookStructure.ReadSheets(package).Where(sheet => sheet.IsWorksheet).ToList();
}
