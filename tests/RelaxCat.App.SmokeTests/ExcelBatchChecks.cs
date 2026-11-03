using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RelaxCat.App;
using RelaxCat.App.Models;

namespace RelaxCat.App.SmokeTests;

/// <summary>Headless checks for the batch workbook page: scan, per-file results, failures and i18n.</summary>
internal static class ExcelBatchChecks
{
    internal static void Run(MainWindow window, Action<Task> complete, Action<string> capture, Action<Func<bool>> pumpUntil)
    {
        var theme = window.Workspace.ThemeIndex;
        var language = window.Workspace.LanguageIndex;
        window.Workspace.ThemeIndex = 1;
        window.Workspace.LanguageIndex = 0;
        Dispatcher.UIThread.RunJobs();

        var page = window.FindControl<RelaxCat.App.Views.ExcelBatchView>("ExcelBatchPage") ??
                   throw new Exception("Missing Excel batch page");
        Click(window.FindControl<Button>("ExcelBatchButton")!);
        pumpUntil(() => page.IsVisible && window.Workspace.ExcelBatchOpen);
        Check(window.Workspace.HomeOpen == false, "Batch navigation did not open its page");

        var batch = window.Workspace.ExcelTool.Batch;
        var input = Path.GetFullPath("artifacts/batch-input");
        var output = Path.GetFullPath("artifacts/batch-output");
        if (Directory.Exists(input)) Directory.Delete(input, true);
        if (Directory.Exists(output)) Directory.Delete(output, true);
        Directory.CreateDirectory(Path.Combine(input, "sub"));
        var sample = Path.GetFullPath("artifacts/excel-tool-sample.xlsx");
        if (File.Exists(sample))
        {
            File.Copy(sample, Path.Combine(input, "a.xlsx"));
            File.Copy(sample, Path.Combine(input, "sub", "b.xlsx"));
        }
        else
        {
            ExcelToolChecks.CreateWorkbook(Path.Combine(input, "a.xlsx"));
            ExcelToolChecks.CreateWorkbook(Path.Combine(input, "sub", "b.xlsx"));
        }
        File.WriteAllText(Path.Combine(input, "bad.xlsx"), "not a zip package");
        File.WriteAllText(Path.Combine(input, "notes.txt"), "ignored");

        batch.InputFolder = input;
        batch.OutputFolder = output;
        batch.Recursive = true;
        batch.CollectUnusedStyles = true;
        Dispatcher.UIThread.RunJobs();
        Check(batch.CanScan && !batch.CanRun, "Batch actions were enabled before scanning");

        complete(window.ScanExcelBatchAsync(input, output, recursive: true));
        pumpUntil(() => batch.HasItems);
        Check(batch.Items.Count == 3, "Batch scan did not find the expected workbooks");
        Check(batch.CanRun, "Batch optimization was not enabled after scanning");
        Check(batch.Items.All(item => item.StatusText == window.Workspace.Text["BatchPending"]),
            "Scanned items are not shown as pending");
        Check(page.FindControl<ItemsControl>("BatchResultsList")!.GetVisualDescendants()
            .OfType<Border>().Count(border => border.Classes.Contains("batchRow")) == 3,
            "Batch result rows were not rendered");
        capture("excel-batch-scan-light.png");

        complete(window.RunExcelBatchAsync());
        pumpUntil(() => batch.Status == window.Workspace.Text["BatchDone"]);
        Check(batch.Summary.Contains("完成 2"), "Batch summary did not count completed files");
        Check(batch.Items.Single(item => item.RelativePath == "bad.xlsx").StatusText == window.Workspace.Text["BatchFailed"],
            "Failed workbook was not reported");
        Check(batch.Items.Single(item => item.RelativePath == "a.xlsx").StatusText == window.Workspace.Text["BatchCompleted"],
            "Completed workbook was not reported");
        var firstOutput = Path.Combine(output, "a-optimized.xlsx");
        var secondOutput = Path.Combine(output, "sub", "b-optimized.xlsx");
        Check(File.Exists(firstOutput) && File.Exists(secondOutput), "Batch outputs were not written");
        Check(!File.Exists(Path.Combine(output, "bad-optimized.xlsx")), "Failed workbook produced an output");
        Check(window.Workspace.Message == batch.Summary, "Batch status bar did not report the final summary");
        capture("excel-batch-result-light.png");

        window.Width = 1000; window.Height = 680; Dispatcher.UIThread.RunJobs();
        Check(page.FindControl<Button>("BatchRunButton")!.IsEffectivelyVisible,
            "Batch actions disappeared at the minimum window size");
        capture("excel-batch-narrow-light.png");
        window.Width = 1280; window.Height = 850; Dispatcher.UIThread.RunJobs();

        window.Workspace.LanguageIndex = 1;
        Dispatcher.UIThread.RunJobs();
        Check(batch.Items.Single(item => item.RelativePath == "bad.xlsx").StatusText == "失敗",
            "Japanese batch status did not refresh");
        Check(batch.Status == "一括最適化が完了しました", "Japanese batch completion label missing");
        window.Workspace.ThemeIndex = 0;
        Dispatcher.UIThread.RunJobs();
        capture("excel-batch-result-dark-ja.png");
        window.Workspace.LanguageIndex = 0;
        window.Workspace.ThemeIndex = 1;
        Dispatcher.UIThread.RunJobs();

        Click(window.FindControl<Button>("HomeButton")!);
        window.Workspace.LanguageIndex = language;
        window.Workspace.ThemeIndex = theme;
        Dispatcher.UIThread.RunJobs();
        Console.WriteLine("PASS: batch folder scan, per-file results, failure continuation and bilingual batch UI.");
    }

    private static void Click(Button button)
    {
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        Dispatcher.UIThread.RunJobs();
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
