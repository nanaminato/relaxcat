using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using System.Diagnostics;
using RelaxCat.App.Models;

namespace RelaxCat.App.Views;

public partial class ExcelToolView : UserControl
{
    public event EventHandler? AnalyzeRequested;
    public event EventHandler? OptimizeRequested;

    public ExcelToolView() => InitializeComponent();

    private void AnalyzeClicked(object? sender, RoutedEventArgs e) => AnalyzeRequested?.Invoke(this, EventArgs.Empty);
    private void OptimizeClicked(object? sender, RoutedEventArgs e) => OptimizeRequested?.Invoke(this, EventArgs.Empty);

    private async void BrowseWorkbookClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model || TopLevel.GetTopLevel(this) is not { } top) return;
        try
        {
            var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = model.Text["WorkbookSection"],
                AllowMultiple = false,
                FileTypeFilter = new[] { new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx", "*.xlsm" } } }
            });
            if (files.Count == 0) return;
            model.ExcelTool.FilePath = files[0].TryGetLocalPath() ?? throw new IOException(model.Text["LocalOnly"]);
        }
        catch (Exception error) { model.Message = model.Text["ExcelAnalysisFailed"] + error.Message; }
    }

    private void SheetClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is WorkspaceModel model && (sender as Control)?.DataContext is ExcelSheetSummary sheet)
            model.ExcelTool.Select(sheet);
    }

    private void OpenOutputFolderClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model) return;
        var path = model.ExcelTool.Result?.OutputPath;
        if (path is null || !File.Exists(path)) return;
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + path + "\"") { UseShellExecute = true });
        }
        catch (Exception error) { model.Message = model.Text["OpenFolderFailed"] + error.Message; }
    }
}
