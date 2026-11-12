using Avalonia.Controls;
using Avalonia.Interactivity;
using RelaxCat.Core.Excel;

namespace RelaxCat.App;

public sealed partial class MainWindow
{
    private void ExcelBatchClicked(object? sender, RoutedEventArgs e) => Workspace.ShowExcelBatch();
    private async void ScanExcelBatchRequested(object? sender, EventArgs e) =>
        await ScanExcelBatchAsync(Workspace.ExcelTool.Batch.InputFolder, Workspace.ExcelTool.Batch.OutputFolder,
            Workspace.ExcelTool.Batch.Recursive);
    private async void RunExcelBatchRequested(object? sender, EventArgs e) => await RunExcelBatchAsync();

    internal async Task ScanExcelBatchAsync(string inputFolder, string outputFolder, bool recursive,
        CancellationToken cancellationToken = default)
    {
        if (Workspace.IsBusy || closed) return;
        if (!Directory.Exists(inputFolder))
        {
            Workspace.Message = Workspace.Text["BatchInputRequired"];
            return;
        }
        operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = operation.Token;
        Workspace.IsBusy = true;
        Workspace.Progress = 0;
        Workspace.Message = Workspace.Text["BatchScanning"];
        try
        {
            var discovery = await Task.Run(() => ExcelWorkbookFinder.Find(inputFolder, recursive,
                string.IsNullOrWhiteSpace(outputFolder) ? null : outputFolder, token), token);
            token.ThrowIfCancellationRequested();
            if (closed) return;
            Workspace.ExcelTool.Batch.ApplyDiscovery(discovery);
            Workspace.Progress = 1;
            Workspace.Message = Workspace.ExcelTool.Batch.HasIssues
                ? Workspace.ExcelTool.Batch.Summary + " · " + Workspace.ExcelTool.Batch.IssuesText
                : Workspace.ExcelTool.Batch.Summary;
        }
        catch (OperationCanceledException) { Workspace.Message = Workspace.Text["BatchScanCancelled"]; }
        catch (Exception error) { Workspace.Message = Workspace.Text["BatchScanFailed"] + error.Message; }
        finally
        {
            operation?.Dispose();
            operation = null;
            Workspace.IsBusy = false;
        }
    }

    internal async Task RunExcelBatchAsync(CancellationToken cancellationToken = default)
    {
        if (Workspace.IsBusy || closed) return;
        var batch = Workspace.ExcelTool.Batch;
        if (!Directory.Exists(batch.InputFolder))
        {
            Workspace.Message = Workspace.Text["BatchInputRequired"];
            return;
        }
        if (string.IsNullOrWhiteSpace(batch.OutputFolder))
        {
            Workspace.Message = Workspace.Text["BatchOutputRequired"];
            return;
        }
        var files = batch.Files.ToArray();
        if (files.Length == 0)
        {
            Workspace.Message = Workspace.Text["BatchNoFiles"];
            return;
        }

        operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = operation.Token;
        Workspace.IsBusy = true;
        Workspace.Progress = 0;
        Workspace.Message = Workspace.Text["BatchStarting"];
        try
        {
            var reporting = new Progress<ExcelBatchProgress>(report =>
            {
                batch.ApplyProgress(report);
                // A completion report already counts the finished file; do not add FileProgress twice.
                var work = report.LastResult is null ? report.Completed + report.FileProgress : report.Completed;
                Workspace.Progress = report.Total > 0 ? Math.Clamp(work / report.Total, 0, 1) : 0;
                var current = Math.Min(report.Completed + (report.LastResult is null ? 1 : 0), report.Total);
                if (report.CurrentFile is { } file)
                    Workspace.Message = string.Format(Workspace.Text["BatchProcessing"], file,
                        Math.Max(1, current), report.Total);
            });
            var options = batch.OptimizationOptions;
            var result = await Task.Run(() => new ExcelBatchOptimizer()
                .Optimize(files, batch.OutputFolder, batch.Overwrite, reporting, token, options, Workspace.Settings.AppendOptimizationSuffix), token);
            if (closed) return;
            batch.Complete(result);
            Workspace.Progress = 1;
            Workspace.Message = result.Cancelled ? Workspace.Text["BatchCancelled"] : batch.Summary;
        }
        catch (OperationCanceledException) { Workspace.Message = Workspace.Text["BatchCancelled"]; }
        catch (Exception error) { Workspace.Message = Workspace.Text["BatchRunFailed"] + error.Message; }
        finally
        {
            operation?.Dispose();
            operation = null;
            Workspace.IsBusy = false;
        }
    }
}
