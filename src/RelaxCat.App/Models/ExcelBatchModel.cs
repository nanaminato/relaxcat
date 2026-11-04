using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia.Media;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.Models;

/// <summary>Batch folder scanner and per-file result list for the Excel optimizer.</summary>
public sealed class ExcelBatchModel : ObservableModel
{
    private readonly UiText text;
    private readonly Dictionary<string, ExcelBatchItemResult> resultMap = new(StringComparer.Ordinal);
    private IReadOnlyList<ExcelBatchFile> discovered = Array.Empty<ExcelBatchFile>();
    private IReadOnlyList<string> lastIssues = Array.Empty<string>();
    private ExcelBatchOutcome? outcome;
    private string inputFolder = "", outputFolder = "", status, summary = "", issuesText = "";
    private bool recursive = true, overwrite;

    internal ExcelBatchModel(UiText text)
    {
        this.text = text;
        status = text["BatchIdle"];
    }

    public ObservableCollection<ExcelBatchItem> Items { get; } = new();
    public IReadOnlyList<ExcelBatchFile> Files => discovered;
    public bool HasItems => Items.Count > 0;
    public bool HasIssues => issuesText.Length > 0;
    public bool CanScan => !string.IsNullOrWhiteSpace(InputFolder) && Directory.Exists(InputFolder);
    public bool CanRun => CanScan && !string.IsNullOrWhiteSpace(OutputFolder) && Items.Count > 0;

    public string InputFolder
    {
        get => inputFolder;
        set
        {
            if (!Set(ref inputFolder, value)) return;
            ResetScan();
            Notify(nameof(CanScan));
            Notify(nameof(CanRun));
        }
    }

    public string OutputFolder
    {
        get => outputFolder;
        set
        {
            if (!Set(ref outputFolder, value)) return;
            ResetResults();
            Notify(nameof(CanRun));
        }
    }

    public bool Recursive
    {
        get => recursive;
        set
        {
            if (!Set(ref recursive, value)) return;
            ResetScan();
            Notify(nameof(CanRun));
        }
    }

    private bool collectUnusedStyles;
    public bool CollectUnusedStyles
    {
        get => collectUnusedStyles;
        set { if (Set(ref collectUnusedStyles, value)) ResetResults(); }
    }
    private bool removeCopiedNamedStyles;
    public bool RemoveCopiedNamedStyles
    {
        get => removeCopiedNamedStyles;
        set { if (Set(ref removeCopiedNamedStyles, value)) ResetResults(); }
    }
    internal ExcelStyleOptimizationOptions OptimizationOptions => new(CollectUnusedStyles: CollectUnusedStyles,
        RemoveCopiedNamedStyles: RemoveCopiedNamedStyles);
    public bool Overwrite { get => overwrite; set => Set(ref overwrite, value); }
    public string Status { get => status; private set => Set(ref status, value); }
    public string Summary { get => summary; private set => Set(ref summary, value); }
    public string IssuesText { get => issuesText; private set => Set(ref issuesText, value); }

    internal void ApplyDiscovery(ExcelBatchDiscovery discovery)
    {
        discovered = discovery.Files;
        lastIssues = discovery.Issues;
        resultMap.Clear();
        outcome = null;
        Items.Clear();
        foreach (var file in discovered) Items.Add(new ExcelBatchItem(file, text));
        IssuesText = lastIssues.Count == 0
            ? ""
            : string.Format(CultureInfo.CurrentCulture, text["BatchIssues"], string.Join("; ", lastIssues));
        Summary = string.Format(CultureInfo.CurrentCulture, text["BatchFound"], discovered.Count);
        Status = discovered.Count == 0 ? text["BatchNoFiles"] : text["BatchReady"];
        Notify(nameof(HasItems));
        Notify(nameof(HasIssues));
        Notify(nameof(CanRun));
        Notify(nameof(Summary));
        Notify(nameof(Status));
    }

    internal void ApplyProgress(ExcelBatchProgress progress)
    {
        if (progress.LastResult is { } result) ApplyResult(result);
        if (progress.CurrentFile is { } file && progress.Total > 0 && progress.Completed < progress.Total)
            Status = string.Format(CultureInfo.CurrentCulture, text["BatchProcessing"], file,
                Math.Min(progress.Completed + 1, progress.Total), progress.Total);
        else if (progress.Completed >= progress.Total) Status = text["BatchDone"];
    }

    internal void ApplyResult(ExcelBatchItemResult result)
    {
        resultMap[result.InputPath] = result;
        Items.FirstOrDefault(item => item.InputPath == result.InputPath)?.Apply(result, text);
        Summary = BuildSummary();
        Notify(nameof(Summary));
        Notify(nameof(CanRun));
    }

    internal void Complete(ExcelBatchOutcome result)
    {
        outcome = result;
        foreach (var item in result.Results) ApplyResult(item);
        Summary = result.Cancelled
            ? string.Format(CultureInfo.CurrentCulture, text["BatchSummaryCancelled"],
                Count(ExcelBatchItemStatus.Completed), Count(ExcelBatchItemStatus.Skipped), Count(ExcelBatchItemStatus.Failed))
            : BuildSummary();
        Status = result.Cancelled ? text["BatchCancelled"] : text["BatchDone"];
        Notify(nameof(Summary));
        Notify(nameof(Status));
    }

    internal void RefreshLanguage()
    {
        foreach (var item in Items) item.RefreshLanguage(text, resultMap.GetValueOrDefault(item.InputPath));
        IssuesText = lastIssues.Count == 0
            ? ""
            : string.Format(CultureInfo.CurrentCulture, text["BatchIssues"], string.Join("; ", lastIssues));
        if (outcome is not null)
        {
            Summary = outcome.Cancelled
                ? string.Format(CultureInfo.CurrentCulture, text["BatchSummaryCancelled"],
                    Count(ExcelBatchItemStatus.Completed), Count(ExcelBatchItemStatus.Skipped), Count(ExcelBatchItemStatus.Failed))
                : BuildSummary();
            Status = outcome.Cancelled ? text["BatchCancelled"] : text["BatchDone"];
        }
        else if (discovered.Count == 0) Status = text["BatchIdle"];
        else if (resultMap.Count > 0) Summary = BuildSummary();
        else Status = text["BatchReady"];
        Notify(nameof(Summary));
        Notify(nameof(Status));
        Notify(nameof(HasIssues));
    }

    private void ResetScan()
    {
        discovered = Array.Empty<ExcelBatchFile>();
        lastIssues = Array.Empty<string>();
        resultMap.Clear();
        outcome = null;
        Items.Clear();
        Summary = "";
        IssuesText = "";
        Status = text["BatchIdle"];
        Notify(nameof(HasItems));
        Notify(nameof(HasIssues));
        Notify(nameof(Summary));
        Notify(nameof(Status));
        Notify(nameof(CanRun));
    }

    private void ResetResults()
    {
        resultMap.Clear();
        outcome = null;
        foreach (var item in Items) item.RefreshLanguage(text, null);
        Summary = discovered.Count == 0 ? "" : string.Format(CultureInfo.CurrentCulture, text["BatchFound"], discovered.Count);
        Status = discovered.Count == 0 ? text["BatchIdle"] : text["BatchReady"];
        Notify(nameof(Summary));
        Notify(nameof(Status));
        Notify(nameof(CanRun));
    }

    private int Count(ExcelBatchItemStatus status) => resultMap.Values.Count(result => result.Status == status);

    private string BuildSummary() => string.Format(CultureInfo.CurrentCulture, text["BatchSummary"],
        Count(ExcelBatchItemStatus.Completed), Count(ExcelBatchItemStatus.Skipped), Count(ExcelBatchItemStatus.Failed));
}

public sealed class ExcelBatchItem : ObservableModel
{
    private static readonly IBrush PendingBrush = new SolidColorBrush(Color.Parse("#8A94A6"));
    private static readonly IBrush CompletedBrush = new SolidColorBrush(Color.Parse("#2E9E5B"));
    private static readonly IBrush SkippedBrush = new SolidColorBrush(Color.Parse("#D19A00"));
    private static readonly IBrush FailedBrush = new SolidColorBrush(Color.Parse("#C0392B"));

    private string statusText = "";
    private string detailText = "";
    private IBrush statusBrush = PendingBrush;

    internal ExcelBatchItem(ExcelBatchFile file, UiText text)
    {
        InputPath = file.Path;
        RelativePath = file.RelativePath;
        SizeText = ExcelToolModel.Bytes(file.SizeBytes);
        RefreshLanguage(text, null);
    }

    public string InputPath { get; }
    public string RelativePath { get; }
    public string SizeText { get; }
    public string StatusText { get => statusText; private set => Set(ref statusText, value); }
    public string DetailText { get => detailText; private set => Set(ref detailText, value); }
    public IBrush StatusBrush { get => statusBrush; private set => Set(ref statusBrush, value); }

    internal void Apply(ExcelBatchItemResult result, UiText text)
    {
        switch (result.Status)
        {
            case ExcelBatchItemStatus.Completed:
                StatusText = text["BatchCompleted"];
                StatusBrush = CompletedBrush;
                var detail = string.Format(CultureInfo.CurrentCulture, text["BatchResultDetail"],
                    ExcelToolModel.Count(result.InputCellXfCount), ExcelToolModel.Count(result.OutputCellXfCount),
                    ExcelToolModel.Bytes(result.InputBytes), ExcelToolModel.Bytes(result.OutputBytes));
                DetailText = result.Verified ? detail + " · " + text["BatchVerified"] : detail;
                if (result.AdvancedOptimizationSkipReason is { } reason)
                    DetailText += " · " + text["AdvancedSkip_" + reason];
                else if (result.TableChanges.Count > 0)
                {
                    DetailText += " · " + string.Join(" · ", result.TableChanges.Where(c => c.Name != "cellXfs")
                        .Select(c => $"{c.Name}: {c.Before:N0} → {c.After:N0}"));
                    var namedStylesRemoved = result.TableChanges.Where(c => c.Name == "cellStyles").Sum(c => c.Before - c.After);
                    if (namedStylesRemoved > 0)
                        DetailText += " · " + text["ConditionalNamedStyleRemovals"] + ": " + ExcelToolModel.Count(namedStylesRemoved);
                }
                break;
            case ExcelBatchItemStatus.Skipped:
                StatusText = text["BatchSkipped"];
                StatusBrush = SkippedBrush;
                DetailText = text["BatchSkippedReason"];
                break;
            case ExcelBatchItemStatus.Cancelled:
                StatusText = text["BatchCancelled"];
                StatusBrush = PendingBrush;
                DetailText = "";
                break;
            default:
                StatusText = text["BatchFailed"];
                StatusBrush = FailedBrush;
                DetailText = result.Error ?? "";
                break;
        }
    }

    internal void RefreshLanguage(UiText text, ExcelBatchItemResult? result)
    {
        if (result is not null) Apply(result, text);
        else
        {
            StatusText = text["BatchPending"];
            StatusBrush = PendingBrush;
            DetailText = SizeText;
        }
    }
}
