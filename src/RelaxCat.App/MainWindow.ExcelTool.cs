using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using RelaxCat.App.Models;
using RelaxCat.Core.Excel;

namespace RelaxCat.App;

public sealed partial class MainWindow
{
    private void ExcelToolClicked(object? sender, RoutedEventArgs e) => Workspace.ShowExcelTool();
    private async void AnalyzeExcelRequested(object? sender, EventArgs e) =>
        await AnalyzeExcelWorkbookAsync(Workspace.ExcelTool.FilePath);
    private async void OptimizeExcelRequested(object? sender, EventArgs e) => await ChooseAndOptimizeExcelAsync();

    private async Task ChooseAndOptimizeExcelAsync()
    {
        if (Workspace.IsBusy || closed) return;
        var input = Workspace.ExcelTool.FilePath;
        if (!ExcelStyleAnalyzer.IsSupported(input) || !File.Exists(input))
        {
            Workspace.Message = Workspace.Text["ExcelNoWorkbook"];
            return;
        }
        var extension = Path.GetExtension(input);
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = Workspace.Text["OptimizeSaveAs"],
            SuggestedFileName = Path.GetFileNameWithoutExtension(input) + "-optimized" + extension,
            DefaultExtension = extension.TrimStart('.'),
            ShowOverwritePrompt = true,
            FileTypeChoices = new[]
            {
                new FilePickerFileType(Workspace.Text["OptimizeSaveAs"]) { Patterns = new[] { "*" + extension } }
            }
        });
        if (closed || file is null) return;
        var output = file.TryGetLocalPath();
        if (output is null)
        {
            Workspace.Message = Workspace.Text["LocalOnly"];
            return;
        }
        await OptimizeExcelWorkbookAsync(input, output);
    }

    internal async Task AnalyzeExcelWorkbookAsync(string path, CancellationToken cancellationToken = default)
    {
        if (Workspace.IsBusy || closed) return;
        if (!ExcelStyleAnalyzer.IsSupported(path) || !File.Exists(path))
        {
            Workspace.Message = Workspace.Text["ExcelNoWorkbook"];
            return;
        }
        var input = Path.GetFullPath(path);
        operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = operation.Token;
        Workspace.IsBusy = true;
        Workspace.Progress = 0;
        Workspace.Message = Workspace.Text["ExcelAnalyzing"];
        try
        {
            var reporting = new Progress<ExcelAnalysisProgress>(report =>
            {
                Workspace.Message = Describe(report);
                Workspace.Progress = AnalysisProgressOf(report);
            });
            var analysis = await Task.Run(() => new ExcelStyleAnalyzer().Analyze(input, token, reporting), token);
            token.ThrowIfCancellationRequested();
            if (closed) return;
            Workspace.ExcelTool.FilePath = input;
            Workspace.ExcelTool.ApplyAnalysis(analysis);
            Workspace.Progress = 1;
            Workspace.Message = analysis.HasDuplicates
                ? string.Format(Workspace.Text["ExcelAnalyzedDuplicates"], ExcelToolModel.Count(analysis.DuplicateCellXfCount))
                : Workspace.Text["ExcelAnalyzed"];
        }
        catch (OperationCanceledException) { Workspace.Message = Workspace.Text["ExcelAnalyzeCancelled"]; }
        catch (Exception error)
        {
            Workspace.ExcelTool.Clear();
            Workspace.Message = Workspace.Text["ExcelAnalysisFailed"] + error.Message;
        }
        finally
        {
            operation?.Dispose();
            operation = null;
            Workspace.IsBusy = false;
        }
    }

    internal async Task OptimizeExcelWorkbookAsync(string inputPath, string outputPath,
        CancellationToken cancellationToken = default)
    {
        if (Workspace.IsBusy || closed) return;
        if (!ExcelStyleAnalyzer.IsSupported(inputPath) || !File.Exists(inputPath))
        {
            Workspace.Message = Workspace.Text["ExcelNoWorkbook"];
            return;
        }
        var input = Path.GetFullPath(inputPath);
        var output = Path.GetFullPath(outputPath);
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
        {
            Workspace.Message = Workspace.Text["ExcelSameFile"];
            return;
        }
        operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = operation.Token;
        Workspace.IsBusy = true;
        Workspace.Progress = 0;
        Workspace.Message = Workspace.Text["ExcelOptimizing"];
        try
        {
            var reporting = new Progress<ExcelOptimizationProgress>(report =>
            {
                Workspace.Message = Describe(report);
                Workspace.Progress = Math.Clamp(ProgressOf(report), 0, .95);
            });
            ExcelOptimizationResult result;
            try
            {
                result = await Task.Run(() => new ExcelStyleOptimizer().Optimize(input, output, reporting, token), token);
            }
            catch (OperationCanceledException) { Workspace.Message = Workspace.Text["ExcelOptimizeCancelled"]; return; }
            catch (Exception error) { Workspace.Message = Workspace.Text["ExcelOptimizeFailed"] + error.Message; return; }

            // The optimized file is already published and verified; cancellation from here only skips the preview refresh.
            Workspace.ExcelTool.FilePath = result.OutputPath;
            Workspace.ExcelTool.ApplyOptimization(result);
            Workspace.Progress = 1;
            Workspace.Message = OptimizedSummary(result);
            try
            {
                // Re-analyze the optimized copy so the page shows the processed worksheet styles, not only numbers.
                var analysis = await Task.Run(() => new ExcelStyleAnalyzer().Analyze(result.OutputPath, token), token);
                token.ThrowIfCancellationRequested();
                if (closed) return;
                Workspace.ExcelTool.ApplyAnalysis(analysis, keepResult: true);
                Workspace.Message = OptimizedSummary(result);
                Workspace.Progress = 1;
            }
            catch (OperationCanceledException) { }
            catch (Exception error) { Workspace.Message = OptimizedSummary(result) + " · " + Workspace.Text["ExcelAnalysisFailed"] + error.Message; }
        }
        finally
        {
            operation?.Dispose();
            operation = null;
            Workspace.IsBusy = false;
        }
    }

    private string OptimizedSummary(ExcelOptimizationResult result) => string.Format(Workspace.Text["ExcelOptimizedSummary"],
        ExcelToolModel.Count(result.InputCellXfCount), ExcelToolModel.Count(result.OutputCellXfCount),
        Path.GetFileName(result.OutputPath));

    private string Describe(ExcelAnalysisProgress report)
    {
        if (report.Phase == ExcelAnalysisPhase.ReadingStyles) return Workspace.Text["ExcelPhaseReadingStyles"];
        if (report.Phase == ExcelAnalysisPhase.Projecting) return Workspace.Text["ExcelPhaseProjection"];
        var phase = Workspace.Text["ExcelPhaseScanning"];
        if (report.Worksheet is not { } sheet) return phase;
        var completed = Math.Min(report.Completed + 1, Math.Max(1, report.Total));
        var message = $"{phase} · {sheet} ({completed}/{report.Total})";
        if (report.RowsScanned > 0)
            message += $" · {ExcelToolModel.Count(report.RowsScanned)} {Workspace.Text["RowsLabel"]}";
        return message;
    }

    private static double AnalysisProgressOf(ExcelAnalysisProgress report)
    {
        if (report.Phase == ExcelAnalysisPhase.ReadingStyles) return .05;
        if (report.Phase == ExcelAnalysisPhase.Projecting) return .92;
        if (report.Total <= 0) return .1;
        var within = report.ExpectedRows > 0
            ? Math.Clamp((double)report.RowsScanned / report.ExpectedRows, 0, 1)
            : 0;
        return .1 + .8 * Math.Clamp((report.Completed + within) / report.Total, 0, 1);
    }

    private string Describe(ExcelOptimizationProgress report)
    {
        var phase = report.Phase switch
        {
            ExcelOptimizationPhase.Reading => Workspace.Text["ExcelPhaseReading"],
            ExcelOptimizationPhase.RewritingWorksheets => Workspace.Text["ExcelPhaseRewriting"],
            ExcelOptimizationPhase.RewritingStyles => Workspace.Text["ExcelPhaseStyles"],
            ExcelOptimizationPhase.Verifying => Workspace.Text["ExcelPhaseVerifying"],
            _ => Workspace.Text["ExcelPhaseWriting"]
        };
        return report.Total > 0 ? phase + $" {report.Completed}/{report.Total}" : phase;
    }

    private static double ProgressOf(ExcelOptimizationProgress report) => report.Phase switch
    {
        ExcelOptimizationPhase.Reading => .1,
        ExcelOptimizationPhase.RewritingWorksheets => .1 + .5 * report.Completed / Math.Max(1, report.Total),
        ExcelOptimizationPhase.RewritingStyles => .64,
        ExcelOptimizationPhase.Verifying => .7 + .2 * report.Completed / Math.Max(1, report.Total),
        ExcelOptimizationPhase.WritingOutput => .95,
        _ => .1
    };
}
