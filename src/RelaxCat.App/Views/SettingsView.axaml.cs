using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using RelaxCat.App.Models;

namespace RelaxCat.App.Views;

public partial class SettingsView : UserControl
{
    public SettingsView() => InitializeComponent();
    private void ValidateClicked(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not WorkspaceModel model) return;
        try { model.Settings.CreateOptions(model.Text); model.Message = model.Text["Valid"]; }
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
