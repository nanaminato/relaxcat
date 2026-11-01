using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using RelaxCat.App.Models;

namespace RelaxCat.App.Views;

public partial class ExcelBatchView : UserControl
{
    public event EventHandler? ScanRequested;
    public event EventHandler? RunRequested;

    public ExcelBatchView() => InitializeComponent();

    private void ScanClicked(object? sender, RoutedEventArgs e) => ScanRequested?.Invoke(this, EventArgs.Empty);
    private void RunClicked(object? sender, RoutedEventArgs e) => RunRequested?.Invoke(this, EventArgs.Empty);

    private async void BrowseInputClicked(object? sender, RoutedEventArgs e) =>
        await BrowseFolderAsync((model, path) => model.ExcelTool.Batch.InputFolder = path);
    private async void BrowseOutputClicked(object? sender, RoutedEventArgs e) =>
        await BrowseFolderAsync((model, path) => model.ExcelTool.Batch.OutputFolder = path);

    private async Task BrowseFolderAsync(Action<WorkspaceModel, string> apply)
    {
        if (DataContext is not WorkspaceModel model || TopLevel.GetTopLevel(this) is not { } top) return;
        try
        {
            var folders = await top.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = model.Text["ExcelBatchTab"],
                AllowMultiple = false
            });
            if (folders.Count == 0) return;
            var path = folders[0].TryGetLocalPath() ?? throw new IOException(model.Text["LocalOnly"]);
            apply(model, path);
        }
        catch (Exception error) { model.Message = model.Text["BatchScanFailed"] + error.Message; }
    }
}
