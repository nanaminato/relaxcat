using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia;
using Avalonia.Layout;
using Avalonia.Media;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.Models;

public sealed record ExcelMetric(string Label, string Value);

public sealed record ExcelComparisonRow(string Label, string Before, string After);

public sealed record ExcelCheckRow(string Label, string Detail, bool Passed);

/// <summary>Presentation state for the Excel style tool. It only holds style metadata, never cell content.</summary>
public sealed class ExcelToolModel : ObservableModel
{
    private readonly UiText text;
    private string filePath = "";
    private ExcelWorkbookSummary? summary;
    private ExcelSheetSummary? selectedSheet;
    private ExcelToolResult? result;
    private ExcelStyleAnalysis? analysis;
    private ExcelOptimizationResult? optimization;

    internal ExcelToolModel(UiText text)
    {
        this.text = text;
        Batch = new ExcelBatchModel(text);
    }

    public ExcelBatchModel Batch { get; }
    private bool collectUnusedStyles;
    public bool CollectUnusedStyles
    {
        get => collectUnusedStyles;
        set { if (Set(ref collectUnusedStyles, value)) Clear(); }
    }
    private bool removeCopiedNamedStyles;
    public bool RemoveCopiedNamedStyles
    {
        get => removeCopiedNamedStyles;
        set { if (Set(ref removeCopiedNamedStyles, value)) Clear(); }
    }
    internal ExcelStyleOptimizationOptions OptimizationOptions => new(CollectUnusedStyles: CollectUnusedStyles,
        RemoveCopiedNamedStyles: RemoveCopiedNamedStyles);

    public string FilePath
    {
        get => filePath;
        set
        {
            if (!Set(ref filePath, value)) return;
            Clear();
            Notify(nameof(FileName));
        }
    }

    public string FileName => string.IsNullOrWhiteSpace(filePath) ? "" : Path.GetFileName(filePath);
    public bool CanAnalyze => !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath) &&
                              ExcelStyleAnalyzer.IsSupported(filePath);
    public bool CanOptimize => CanAnalyze && analysis?.HasStyleChanges == true;
    public bool HasSummary => summary is not null;
    public bool HasResult => result is not null;
    public bool HasSheets => Sheets.Count > 0;
    public bool HasSelectedSheet => selectedSheet is not null;

    public ExcelWorkbookSummary? Summary => summary;
    public ExcelToolResult? Result => result;
    public ObservableCollection<ExcelSheetSummary> Sheets { get; } = new();

    public ExcelSheetSummary? SelectedSheet
    {
        get => selectedSheet;
        private set
        {
            if (Set(ref selectedSheet, value)) Notify(nameof(HasSelectedSheet));
        }
    }

    public void Select(ExcelSheetSummary sheet)
    {
        if (!Sheets.Contains(sheet)) return;
        foreach (var candidate in Sheets) candidate.IsSelected = ReferenceEquals(candidate, sheet);
        SelectedSheet = sheet;
    }

    internal void ApplyAnalysis(ExcelStyleAnalysis value, bool keepResult = false)
    {
        if (!keepResult) ResetResult();
        analysis = value;
        summary = BuildSummary(value, optimization);
        RebuildSheets(value, SelectedSheet?.Name);
        Notify(nameof(Summary));
        Notify(nameof(HasSummary));
        Notify(nameof(CanAnalyze));
        Notify(nameof(CanOptimize));
    }

    internal void ApplyOptimization(ExcelOptimizationResult value)
    {
        optimization = value;
        result = BuildResult(value);
        Notify(nameof(Result));
        Notify(nameof(HasResult));
    }

    internal void Clear()
    {
        analysis = null;
        summary = null;
        optimization = null;
        result = null;
        Sheets.Clear();
        SelectedSheet = null;
        Notify(nameof(Summary));
        Notify(nameof(HasSummary));
        Notify(nameof(Result));
        Notify(nameof(HasResult));
        Notify(nameof(HasSheets));
        Notify(nameof(CanAnalyze));
        Notify(nameof(CanOptimize));
    }

    internal void RefreshLanguage()
    {
        Batch.RefreshLanguage();
        if (analysis is null) return;
        var selected = SelectedSheet?.Name;
        summary = BuildSummary(analysis, optimization);
        RebuildSheets(analysis, selected);
        if (optimization is not null) result = BuildResult(optimization);
        Notify(nameof(Summary));
        Notify(nameof(HasSummary));
        Notify(nameof(Result));
        Notify(nameof(HasResult));
        Notify(nameof(HasSheets));
    }

    private void ResetResult()
    {
        optimization = null;
        result = null;
        Notify(nameof(Result));
        Notify(nameof(HasResult));
    }

    private void RebuildSheets(ExcelStyleAnalysis value, string? preferred)
    {
        Sheets.Clear();
        foreach (var sheet in value.Worksheets) Sheets.Add(new ExcelSheetSummary(sheet, text));
        var target = Sheets.FirstOrDefault(sheet => sheet.Name == preferred) ?? Sheets.FirstOrDefault();
        if (target is not null) Select(target);
        else SelectedSheet = null;
        Notify(nameof(HasSheets));
    }

    private ExcelWorkbookSummary BuildSummary(ExcelStyleAnalysis value, ExcelOptimizationResult? optimized)
    {
        var comparison = optimized is null
            ? new List<ExcelComparisonRow>
            {
                new(text["RowCellXfs"], Count(value.CellXfCount), Count(value.ProjectedCellXfCount)),
                new("cellStyleXfs", Count(value.CellStyleXfCount), Count(value.ProjectedCellStyleXfCount)),
                new(text["StylesPart"], Bytes(value.StylesXmlBytes), Bytes(value.ProjectedStylesXmlBytes))
            }
            : new List<ExcelComparisonRow>
            {
                new(text["RowCellXfs"], Count(optimized.InputCellXfCount), Count(optimized.OutputCellXfCount)),
                new("cellStyleXfs", Count(optimized.InputCellStyleXfCount), Count(optimized.OutputCellStyleXfCount)),
                new(text["StylesPart"], Bytes(optimized.InputStylesXmlBytes), Bytes(optimized.OutputStylesXmlBytes))
            };
        return new ExcelWorkbookSummary(
            value.Path,
            Path.GetFileName(value.Path),
            value.MacroEnabled ? text["MacroWorkbook"] : "",
            text["StyleDataReduction"],
            optimized is null ? text["AfterColumn"] : text["AfterColumnActual"],
            Percent(optimized?.StylesXmlReductionRatio ?? (value.StylesXmlBytes == 0 ? 0 :
                (double)(value.StylesXmlBytes - value.ProjectedStylesXmlBytes) / value.StylesXmlBytes)),
            new List<ExcelMetric>
            {
                new(text["FileSize"], Bytes(value.FileSize)),
                new(text["StylesPart"], Bytes(value.StylesXmlBytes)),
                new(text["StylesCompressed"], Bytes(value.StylesXmlCompressedBytes)),
                new(text["UniqueBaseFormats"], Count(value.UniqueCellStyleXfCount)),
                new(text["BaseDeduplication"], value.CanDeduplicateBaseStyles ? text["BaseDedupAvailable"] : text["BaseDedupSkipped"]),
                new(text["BaseReferences"], Count(value.BaseStyleReferences.Count(r => r.CellXfReferences + r.NamedStyleReferences > 0))),
                new(text["CellXfs"], Count(value.CellXfCount)),
                new(text["UniqueFormats"], Count(value.UniqueCellXfCount)),
                new(text["DuplicateFormats"], Count(value.DuplicateCellXfCount)),
                new(text["UnreferencedFormats"], Count(value.UnreferencedCellXfCount)),
                new(text["StyleTableParts"], $"{Count(value.FontCount)} · {Count(value.FillCount)} · {Count(value.BorderCount)}"),
                new(text["WorksheetCount"], Count(value.SheetCount))
            },
            comparison)
        {
            Diagnostics = value.Sections.Select(section => new ExcelMetric(section.Name,
                $"{Count(section.Count)} · {Bytes(section.SerializedBytes)} · {Percent(section.SerializedShare)}\n" +
                $"{text["DuplicateFormats"]}: {Count(section.DuplicateCount)} · {text["KnownReferences"]}: {section.ReferencedCount?.ToString("N0") ?? "?"} · {text["UnreferencedFormats"]}: {section.UnreferencedCount?.ToString("N0") ?? "?"}"
            )).Concat(value.TableChanges.Select(change => new ExcelMetric(change.Name + " · " + text["TableChange"],
                $"{Count(change.Before)} → {Count(change.After)} · {text[change.Name == "cellStyles" ? "ConditionalNamedStyleRemovals" : "DuplicateFormats"]}: {Count(change.DuplicatesRemoved)} · {text["UnusedRemoved"]}: {Count(change.UnusedRemoved)}")))
                .Append(new ExcelMetric(text["ConditionalNamedStyleRemovals"],
                    Count(value.TableChanges.Where(change => change.Name == "cellStyles").Sum(change => change.Before - change.After))))
                .Concat(BuildNamedStyleCandidates(value.NamedStyleCleanupCandidates))
                .Append(new ExcelMetric(text["AdvancedOptimization"], value.AdvancedOptimizationSkipReason is null
                    ? text["AdvancedAvailable"] : text["AdvancedSkip_" + value.AdvancedOptimizationSkipReason]))
                .Concat(BuildExtensionDiagnostics(value.ExtensionDiagnostics)).ToList()
        };
    }

    private IEnumerable<ExcelMetric> BuildNamedStyleCandidates(IReadOnlyList<string> candidates)
    {
        if (candidates.Count == 0) yield break;
        const int displayLimit = 100;
        yield return new ExcelMetric(text["NamedStyleCleanupCandidates"],
            string.Format(CultureInfo.CurrentCulture, text["NamedStyleCandidateCount"],
                Count(candidates.Count), Count(Math.Min(candidates.Count, displayLimit))) + "\n" +
            string.Join("\n", candidates.Take(displayLimit)));
    }

    private IEnumerable<ExcelMetric> BuildExtensionDiagnostics(ExtensionTypeDiagnostics diagnostics)
    {
        yield return new ExcelMetric(text["ExtensionDiagnosticsTitle"], text["ExtensionDiagnosticsHint"]);
        yield return new ExcelMetric(text["ExtensionDiagnosticsScope"], text["ExtensionDiagnosticsScopeHint"]);
        if (diagnostics.Truncated)
            yield return new ExcelMetric(text["ExtensionDiagnosticsLimit"], text["ExtensionDiagnosticsTruncated"]);
        if (diagnostics.Entries.Count == 0)
            yield return new ExcelMetric(text["ExtensionDiagnosticsTypes"], text["ExtensionDiagnosticsEmpty"]);
        foreach (var entry in diagnostics.Entries)
        {
            var key = entry.Kind switch
            {
                "uri" => "ExtensionDiagnosticUri",
                "namespace" => "ExtensionDiagnosticNamespace",
                "element" => "ExtensionDiagnosticElement",
                "attribute" => "ExtensionDiagnosticAttribute",
                _ => "ExtensionDiagnosticsTypes"
            };
            yield return new ExcelMetric(text[key] + " · " +
                string.Format(CultureInfo.CurrentCulture, text["ExtensionDiagnosticCount"], Count(entry.Count)), entry.Name);
        }
    }

    private ExcelToolResult BuildResult(ExcelOptimizationResult value)
    {
        var checks = value.Checks
            .Select(check => new ExcelCheckRow(text[CheckKey(check.Key)], check.Detail, check.Passed))
            .ToList();
        return new ExcelToolResult(value.OutputPath, Path.GetFileName(value.OutputPath),
            Path.GetDirectoryName(value.OutputPath) ?? "",
            new List<ExcelMetric>
            {
                new(text["FileSize"], Bytes(value.InputFileSize) + " → " + Bytes(value.OutputFileSize)),
                new(text["CellXfs"], Count(value.InputCellXfCount) + " → " + Count(value.OutputCellXfCount)),
                new(text["StylesPart"], Bytes(value.InputStylesXmlBytes) + " → " + Bytes(value.OutputStylesXmlBytes)),
                new("cellStyleXfs", Count(value.InputCellStyleXfCount) + " → " + Count(value.OutputCellStyleXfCount)),
                new(text["StyleDataReduction"], Percent(value.StylesXmlReductionRatio)),
                new(text["UnusedRemoved"], Count(value.TableChanges.Sum(change => change.UnusedRemoved))),
                new(text["ConditionalNamedStyleRemovals"],
                    Count(value.TableChanges.Where(change => change.Name == "cellStyles").Sum(change => change.Before - change.After)))
            }.Concat(value.TableChanges.Where(change => change.Name is not ("cellXfs" or "cellStyleXfs"))
                .Select(change => new ExcelMetric(change.Name, Count(change.Before) + " → " + Count(change.After))))
             .Append(new ExcelMetric(text["AdvancedOptimization"], value.AdvancedOptimizationSkipReason is null
                 ? text["AdvancedAvailable"] : text["AdvancedSkip_" + value.AdvancedOptimizationSkipReason])).ToList(),
            checks,
            value.Verified ? string.Format(CultureInfo.CurrentCulture, text["VerificationPassed"], checks.Count) : text["VerificationFailed"],
            value.Verified);
    }

    private static string CheckKey(string key) => key switch
    {
        "stylesSections" => "CheckStylesSections",
        "cellXfs" => "CheckCellXfs",
        "references" => "CheckReferences",
        "parts" => "CheckParts",
        _ => "CheckPackage"
    };

    internal static string Count(int value) => value.ToString("N0", CultureInfo.CurrentCulture);
    internal static string Percent(double ratio) => (ratio * 100).ToString("0.##", CultureInfo.CurrentCulture) + "%";

    internal static string Bytes(long value) => value switch
    {
        >= 1024L * 1024 * 1024 => (value / (1024d * 1024 * 1024)).ToString("0.##", CultureInfo.CurrentCulture) + " GB",
        >= 1024L * 1024 => (value / (1024d * 1024)).ToString("0.##", CultureInfo.CurrentCulture) + " MB",
        >= 1024 => (value / 1024d).ToString("0.#", CultureInfo.CurrentCulture) + " KB",
        _ => value.ToString("N0", CultureInfo.CurrentCulture) + " B"
    };
}

public sealed class ExcelWorkbookSummary
{
    internal ExcelWorkbookSummary(string path, string fileName, string macroText, string reductionLabel,
        string afterLabel, string reductionText, IReadOnlyList<ExcelMetric> metrics, IReadOnlyList<ExcelComparisonRow> comparison)
    {
        Path = path;
        FileName = fileName;
        MacroText = macroText;
        ReductionLabel = reductionLabel;
        AfterLabel = afterLabel;
        ReductionText = reductionText;
        Metrics = metrics;
        Comparison = comparison;
    }

    public string Path { get; }
    public string FileName { get; }
    public string MacroText { get; }
    public bool HasMacro => MacroText.Length > 0;
    public string ReductionLabel { get; }
    public string AfterLabel { get; }
    public string ReductionText { get; }
    public IReadOnlyList<ExcelMetric> Metrics { get; }
    public IReadOnlyList<ExcelComparisonRow> Comparison { get; }
    public IReadOnlyList<ExcelMetric> Diagnostics { get; init; } = Array.Empty<ExcelMetric>();
}

public sealed class ExcelSheetSummary : ObservableModel
{
    private bool isSelected;

    internal ExcelSheetSummary(WorksheetStyleSummary sheet, UiText text)
    {
        Name = sheet.Name;
        Hidden = sheet.Hidden;
        HiddenText = text["SheetHidden"];
        UsedRangeText = string.IsNullOrEmpty(sheet.UsedRange) ? "—" : sheet.UsedRange;
        CellCountText = ExcelToolModel.Count(sheet.CellCount);
        StyleCountText = ExcelToolModel.Count(sheet.DistinctStyleCount);
        HasMore = sheet.HiddenStyleCount > 0;
        MoreText = HasMore ? string.Format(CultureInfo.CurrentCulture, text["MoreStyles"], sheet.HiddenStyleCount) : "";
        EmptyText = text["NoStylesInSheet"];
        foreach (var usage in sheet.Styles) Styles.Add(new ExcelStyleSwatch(usage, text));
    }

    public string Name { get; }
    public bool Hidden { get; }
    public string HiddenText { get; }
    public string UsedRangeText { get; }
    public string CellCountText { get; }
    public string StyleCountText { get; }
    public bool HasMore { get; }
    public string MoreText { get; }
    public string EmptyText { get; }
    public bool HasStyles => Styles.Count > 0;
    public ObservableCollection<ExcelStyleSwatch> Styles { get; } = new();

    public bool IsSelected { get => isSelected; set => Set(ref isSelected, value); }
}

public sealed class ExcelStyleSwatch
{
    internal ExcelStyleSwatch(StyleUsage usage, UiText text)
    {
        var style = usage.Style;
        Fill = Brush(style.FillColor);
        BorderBrush = Brush(style.BorderColor);
        var width = style.Borders == BorderSides.None ? 0 : Math.Clamp(style.BorderThickness, 1, 3);
        BorderThickness = new Thickness(
            style.Borders.HasFlag(BorderSides.Left) ? width : 0,
            style.Borders.HasFlag(BorderSides.Top) ? width : 0,
            style.Borders.HasFlag(BorderSides.Right) ? width : 0,
            style.Borders.HasFlag(BorderSides.Bottom) ? width : 0);
        Foreground = Brush(style.FontColor) ?? Brushes.Black;
        FontFamily = new FontFamily(style.FontName);
        FontSize = Math.Clamp(style.FontSize, 8, 26);
        FontWeight = style.Bold ? FontWeight.Bold : FontWeight.Normal;
        FontStyle = style.Italic ? FontStyle.Italic : FontStyle.Normal;
        HorizontalTextAlignment = style.HorizontalAlignment switch
        {
            PreviewHorizontalAlignment.Center or PreviewHorizontalAlignment.CenterContinuous => TextAlignment.Center,
            PreviewHorizontalAlignment.Right => TextAlignment.Right,
            PreviewHorizontalAlignment.Fill or PreviewHorizontalAlignment.Justify or PreviewHorizontalAlignment.Distributed => TextAlignment.Justify,
            _ => TextAlignment.Left
        };
        VerticalContentAlignment = style.VerticalAlignment switch
        {
            PreviewVerticalAlignment.Top => VerticalAlignment.Top,
            PreviewVerticalAlignment.Center or PreviewVerticalAlignment.Justify or PreviewVerticalAlignment.Distributed => VerticalAlignment.Center,
            _ => VerticalAlignment.Bottom
        };
        Name = style.IsDefault ? text["NormalFormat"] : string.Format(CultureInfo.CurrentCulture, text["FormatLabel"], style.StyleId + 1);
        UsageText = string.Format(CultureInfo.CurrentCulture, text["CellUsage"], ExcelToolModel.Count(usage.CellCount));
        Details = BuildDetails(style, text);
    }

    public IBrush Fill { get; }
    public IBrush BorderBrush { get; }
    public Thickness BorderThickness { get; }
    public IBrush Foreground { get; }
    public FontFamily FontFamily { get; }
    public double FontSize { get; }
    public FontWeight FontWeight { get; }
    public FontStyle FontStyle { get; }
    public TextAlignment HorizontalTextAlignment { get; }
    public VerticalAlignment VerticalContentAlignment { get; }
    public string Sample => "Aa";
    public string Name { get; }
    public string UsageText { get; }
    public string Details { get; }

    private static IBrush Brush(string? hex) =>
        hex is null ? Brushes.Transparent : new SolidColorBrush(Color.Parse(hex));

    private static string BuildDetails(CellStylePreview style, UiText text)
    {
        var emphasis = new List<string>();
        if (style.Bold) emphasis.Add(text["BoldLabel"]);
        if (style.Italic) emphasis.Add(text["ItalicLabel"]);
        if (style.Underline) emphasis.Add(text["UnderlineLabel"]);
        if (style.Strike) emphasis.Add(text["StrikeLabel"]);
        var font = style.FontName + " · " + style.FontSize.ToString("0.#", CultureInfo.CurrentCulture) + " pt";
        if (emphasis.Count > 0) font += " · " + string.Join(" · ", emphasis);

        var sides = new List<string>();
        if (style.Borders.HasFlag(BorderSides.Left)) sides.Add(text["LeftSide"]);
        if (style.Borders.HasFlag(BorderSides.Right)) sides.Add(text["RightSide"]);
        if (style.Borders.HasFlag(BorderSides.Top)) sides.Add(text["TopSide"]);
        if (style.Borders.HasFlag(BorderSides.Bottom)) sides.Add(text["BottomSide"]);
        var border = sides.Count == 0
            ? text["NoBorder"]
            : string.Join(" ", sides) + " · " + style.BorderThickness.ToString("0.#", CultureInfo.CurrentCulture) + " pt";

        return string.Join('\n',
            text["FontLabel"] + "：" + font,
            text["NumberFormatLabel"] + "：" + style.NumberFormat,
            text["BorderLabel"] + "：" + border,
            text["AlignmentLabel"] + "：" + Horizontal(style.HorizontalAlignment, text) + " / " + Vertical(style.VerticalAlignment, text),
            text["FillLabel"] + "：" + (style.FillColor ?? text["NoFill"]));
    }

    private static string Horizontal(PreviewHorizontalAlignment alignment, UiText text) => alignment switch
    {
        PreviewHorizontalAlignment.Left => text["AlignLeft"],
        PreviewHorizontalAlignment.Center => text["AlignCenter"],
        PreviewHorizontalAlignment.Right => text["AlignRight"],
        PreviewHorizontalAlignment.Fill => text["AlignFill"],
        PreviewHorizontalAlignment.Justify => text["AlignJustify"],
        PreviewHorizontalAlignment.CenterContinuous => text["AlignCenterContinuous"],
        PreviewHorizontalAlignment.Distributed => text["AlignDistributed"],
        _ => text["AlignGeneral"]
    };

    private static string Vertical(PreviewVerticalAlignment alignment, UiText text) => alignment switch
    {
        PreviewVerticalAlignment.Top => text["AlignTop"],
        PreviewVerticalAlignment.Center => text["AlignCenter"],
        PreviewVerticalAlignment.Justify => text["AlignJustifyVertical"],
        PreviewVerticalAlignment.Distributed => text["AlignDistributedVertical"],
        _ => text["AlignBottom"]
    };
}

public sealed class ExcelToolResult
{
    internal ExcelToolResult(string outputPath, string outputName, string outputDirectory,
        IReadOnlyList<ExcelMetric> metrics, IReadOnlyList<ExcelCheckRow> checks, string verificationText, bool verified)
    {
        OutputPath = outputPath;
        OutputName = outputName;
        OutputDirectory = outputDirectory;
        Metrics = metrics;
        Checks = checks;
        VerificationText = verificationText;
        Verified = verified;
    }

    public string OutputPath { get; }
    public string OutputName { get; }
    public string OutputDirectory { get; }
    public IReadOnlyList<ExcelMetric> Metrics { get; }
    public IReadOnlyList<ExcelCheckRow> Checks { get; }
    public string VerificationText { get; }
    public bool Verified { get; }
}
