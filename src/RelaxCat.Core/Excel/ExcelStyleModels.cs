namespace RelaxCat.Core.Excel;

public enum PreviewHorizontalAlignment { General, Left, Center, Right, Fill, Justify, CenterContinuous, Distributed }

public enum PreviewVerticalAlignment { Bottom, Top, Center, Justify, Distributed }

[Flags]
public enum BorderSides { None = 0, Left = 1, Right = 2, Top = 4, Bottom = 8 }

/// <summary>
/// A read-only description of one cell format as it would be displayed, built from
/// styles.xml plus the workbook theme. It never contains cell values or formulas.
/// </summary>
public sealed record CellStylePreview(
    int StyleId,
    bool IsDefault,
    string? FillColor,
    string FillPattern,
    string FontName,
    double FontSize,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    string FontColor,
    BorderSides Borders,
    string? BorderColor,
    double BorderThickness,
    PreviewHorizontalAlignment HorizontalAlignment,
    PreviewVerticalAlignment VerticalAlignment,
    bool WrapText,
    string NumberFormat);

public sealed record StyleUsage(int StyleId, int CellCount, CellStylePreview Style);

public sealed record WorksheetStyleSummary(
    string Name,
    bool Hidden,
    string UsedRange,
    int CellCount,
    int DistinctStyleCount,
    int HiddenStyleCount,
    IReadOnlyList<StyleUsage> Styles);

/// <summary>A read-only analysis of an .xlsx/.xlsm package. The source workbook is never modified.</summary>
public sealed record ExcelStyleAnalysis(
    string Path,
    long FileSize,
    bool MacroEnabled,
    int FontCount,
    int FillCount,
    int BorderCount,
    int CellStyleXfCount,
    int CellXfCount,
    int UniqueCellXfCount,
    int ReferencedCellXfCount,
    long StylesXmlBytes,
    long ProjectedStylesXmlBytes,
    IReadOnlyList<WorksheetStyleSummary> Worksheets)
{
    public bool IsLossy { get; init; }
    public string VerificationScope { get; init; } = "SemanticEquivalenceAndPackageIntegrity";
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public IReadOnlyList<StyleTableChange> TableChanges { get; init; } = Array.Empty<StyleTableChange>();
    public string? AdvancedOptimizationSkipReason { get; init; }
    public IReadOnlyList<string> AdvancedOptimizationSkipReasons { get; init; } = [];
    public IReadOnlyList<UnsupportedPartDiagnostic> UnsupportedPartDetails { get; init; } = [];
    public IReadOnlyList<string> RetainedStyleTables { get; init; } = [];
    public IReadOnlyList<StyleReferenceEdge> ReferenceGraph { get; init; } = [];
    public ExtensionTypeDiagnostics ExtensionDiagnostics { get; init; } = new([], false);
    public IReadOnlyList<string> NamedStyleCleanupCandidates { get; init; } = [];
    /// <summary>Coverage of remappable cell/base/component indices, not inline differential formats.</summary>
    public bool ReferenceGraphComplete => AdvancedOptimizationSkipReason is null;
    public long ProjectedStylesXmlCompressedBytesEstimate { get; init; }
    public int ProjectedCellXfCount { get; init; }
    public int ProjectedCellStyleXfCount { get; init; }
    public bool HasStyleChanges => TableChanges.Count > 0;
    public int SheetCount => Worksheets.Count;
    public int DuplicateCellXfCount => CellXfCount - UniqueCellXfCount;
    public int UnreferencedCellXfCount => Math.Max(0, CellXfCount - ReferencedCellXfCount);
    public double ReductionRatio => CellXfCount <= 0 ? 0 : (double)DuplicateCellXfCount / CellXfCount;
    public long StylesXmlCompressedBytes { get; init; }
    public IReadOnlyList<StyleSectionDiagnostic> Sections { get; init; } = Array.Empty<StyleSectionDiagnostic>();
    public IReadOnlyList<BaseStyleReference> BaseStyleReferences { get; init; } = Array.Empty<BaseStyleReference>();
    public int UniqueCellStyleXfCount { get; init; }
    public bool CanDeduplicateBaseStyles { get; init; }
    public bool HasDuplicates => DuplicateCellXfCount > 0 ||
        (CanDeduplicateBaseStyles && UniqueCellStyleXfCount < CellStyleXfCount);
    public bool IsExcelFormatLimitRisk => UniqueCellXfCount > 64000;
}

public enum ExcelOptimizationPhase { Reading, RewritingStyles, RewritingWorksheets, Verifying, WritingOutput }

public sealed record ExcelOptimizationProgress(ExcelOptimizationPhase Phase, int Completed, int Total);

public enum ExcelAnalysisPhase { ReadingStyles, ScanningWorksheets, Projecting }

/// <summary>
/// Analysis progress. <paramref name="Completed"/> counts finished worksheets while
/// <paramref name="RowsScanned"/> reports rows inside the current worksheet, so a single
/// large sheet still produces visible progress.
/// </summary>
public sealed record ExcelAnalysisProgress(
    ExcelAnalysisPhase Phase,
    int Completed,
    int Total,
    string? Worksheet = null,
    int RowsScanned = 0,
    int ExpectedRows = 0);

/// <summary>A verification check. <paramref name="Key"/> is a stable token localized by the UI.</summary>
public sealed record OptimizationCheck(string Key, bool Passed, string Detail);

public sealed record ExcelOptimizationResult(
    string InputPath,
    string OutputPath,
    long InputFileSize,
    long OutputFileSize,
    int InputCellXfCount,
    int OutputCellXfCount,
    long InputStylesXmlBytes,
    long OutputStylesXmlBytes,
    int RemappedReferences,
    int WorksheetCount,
    IReadOnlyList<OptimizationCheck> Checks,
    bool Verified)
{
    /// <summary>True for the explicitly destructive policy, even when a structural gate prevents edits.</summary>
    public bool IsLossy { get; init; }
    public string VerificationScope { get; init; } = "SemanticEquivalenceAndPackageIntegrity";
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public IReadOnlyList<StyleTableChange> TableChanges { get; init; } = Array.Empty<StyleTableChange>();
    public string? AdvancedOptimizationSkipReason { get; init; }
    /// <summary>Complete original-to-output mappings; -1 denotes an opt-in GC removal.</summary>
    public IReadOnlyList<int> CellXfMap { get; init; } = [];
    public IReadOnlyList<int> CellStyleXfMap { get; init; } = [];
    public IReadOnlyList<int> NamedStyleMap { get; init; } = [];
    public int InputCellStyleXfCount { get; init; }
    public int OutputCellStyleXfCount { get; init; }
    public double StylesXmlReductionRatio => InputStylesXmlBytes <= 0 ? 0 :
        (double)(InputStylesXmlBytes - OutputStylesXmlBytes) / InputStylesXmlBytes;
    public int RemovedCellXfCount => InputCellXfCount - OutputCellXfCount;
    public long SizeDelta => InputFileSize - OutputFileSize;
    public double ReductionRatio => InputCellXfCount <= 0 ? 0 : (double)RemovedCellXfCount / InputCellXfCount;
}

/// <summary>Section bytes are standalone UTF-8 serialization sizes, not ZIP sizes or exact source spans.
/// Null reference counts mean coverage is unknown; unreferenced never implies safe to delete.</summary>
public sealed record StyleSectionDiagnostic(string Name, int Count, long SerializedBytes,
    double SerializedShare, int DuplicateCount, int? ReferencedCount, int? UnreferencedCount)
{
    /// <summary>Includes exact duplicates. Equivalence preserves apply flags and inheritance.</summary>
    public int SemanticDuplicateCount { get; init; }
    public int SafeDuplicateRemovalCount { get; init; }
    public int SafeUnusedRemovalCount { get; init; }
    public int ConditionalNamedStyleRemovalCount { get; init; }
}
/// <summary>Aggregated worksheet references or an individual table edge. No cell contents.
/// TargetIndex is a zero-based table index except numFmts/builtinNumberFormats, where it is the numFmtId.</summary>
public sealed record StyleReferenceEdge(string Source, int? SourceIndex, string Target, int TargetIndex,
    string Kind, int Count = 1);
public sealed record BaseStyleReference(int Index, int CellXfReferences, int NamedStyleReferences);

/// <summary>Exact component merging is enabled by default; unreachable-definition removal is opt-in.</summary>
public sealed record ExcelStyleOptimizationOptions(bool DeduplicateComponents = true, bool CollectUnusedStyles = false,
    bool RemoveCopiedNamedStyles = false, bool DestructiveStyleCleanup = false)
{
    public const string DestructiveStyleRisk = "DESTRUCTIVE / LOSSY: discard all named styles except Normal, " +
        "disconnect named-style inheritance, force direct cell formatting, and collect unused definitions. " +
        "Formatting, future edits, protection, macros and external name/index-based automation may change or break. " +
        "VBA bytes are retained, not macro behavior. Verification checks the chosen rewrite, known references and " +
        "package contents only; it does not prove visual equivalence or that Microsoft Excel opens without repair. " +
        "Unknown structural dependencies still block rewriting. Use disposable copies and retain the original.";
    public const string DestructiveVerificationScope = "KnownStructureAndContentOnly_NotVisualOrSemanticEquivalence";
    public const string CopiedNamedStyleRisk = "Copied-name cleanup may remove names used by external programs or templates. " +
        "A suffix is only a candidate heuristic, not proof of disuse. Built-ins, metadata-bearing styles and names referenced by any retained cell format are retained; " +
        "VBA, extensions and unknown dependencies disable cleanup. Review the analysis and retain the original workbook.";
}
public sealed record StyleTableChange(string Name, int Before, int After, int DuplicatesRemoved, int UnusedRemoved)
{
    public int DestructiveRemoved { get; init; }
}
