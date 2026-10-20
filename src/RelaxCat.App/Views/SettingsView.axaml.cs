using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using RelaxCat.App.Models;

namespace RelaxCat.App.Views;

public partial class SettingsView : UserControl
{
    public event EventHandler? ApplyCaptureRequested;
    public event EventHandler? StopCaptureRequested;
    public SettingsView() => InitializeComponent();
    public void OpenCapturePage() => CaptureTab.IsSelected = true;
    private void ApplyCaptureClicked(object? sender, RoutedEventArgs e) => ApplyCaptureRequested?.Invoke(this, EventArgs.Empty);
    private void StopCaptureClicked(object? sender, RoutedEventArgs e) => StopCaptureRequested?.Invoke(this, EventArgs.Empty);
    private async void BrowseCaptureFolderClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model || TopLevel.GetTopLevel(this) is not { } top) return;
        try
        {
            var folders = await top.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = model.Text["CaptureFolder"], AllowMultiple = false });
            if (folders.Count > 0 && model.CanApplyCapture)
                model.Capture.Folder = folders[0].TryGetLocalPath() ?? throw new IOException(model.Text["LocalOnly"]);
        }
        catch (Exception error) { model.Message = model.Text["CaptureError"] + error.Message; }
    }
    private void ValidateClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model) return;
        try
        {
            if (CaptureTab.IsSelected) { model.Capture.CreateOptions(model.Text); model.Message = model.Text["CaptureValid"]; }
            else { model.Settings.CreateOptions(model.Text); model.Message = model.Text["Valid"]; }
        }
        catch (Exception error) { model.Message = model.Text["Invalid"] + error.Message; }
    }
    private async void BrowseTemplateClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model || TopLevel.GetTopLevel(this) is not { } top) return;
        try
        {
            var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = model.Text["Template"], AllowMultiple = false,
                FileTypeFilter = new[] { new FilePickerFileType("Excel") { Patterns = new[] { "*.xltx", "*.xltm", "*.xlt", "*.xlsx", "*.xls" } } }
            });
            if (files.Count > 0 && model.IsIdle)
                model.Settings.TemplatePath = files[0].TryGetLocalPath() ?? throw new IOException(model.Text["LocalOnly"]);
        }
        catch (Exception error) { model.Message = model.Text["Invalid"] + error.Message; }
    }
}
