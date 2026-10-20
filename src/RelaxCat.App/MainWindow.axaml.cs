using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using RelaxCat.App.Models;
using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;

namespace RelaxCat.App;

public sealed partial class MainWindow : Window
{
    internal WorkspaceModel Workspace { get; }
    private readonly ExcelSession excel = new();
    private readonly ImageExporter exporter = new();
    private CancellationTokenSource? operation;
    private Bitmap? clipboardBitmap;
    private int previewGeneration;
    private Control? previewInvoker;
    private bool closed;

    public MainWindow() : this(null) { }
    internal MainWindow(SettingsStore? store)
    {
        Workspace = new WorkspaceModel(store);
        InitializeComponent();
        DataContext = Workspace;
        Opened += async (_, _) => { if (Workspace.Capture.ProviderIndex != 0) await ApplyCaptureAsync(); };
        Closing += WindowClosing;
        Closed += (_, _) =>
        {
            closed = true; previewGeneration++;
            Workspace.Dispose(); clipboardBitmap?.Dispose(); excel.Dispose();
        };
    }

    private void HomeClicked(object? sender, RoutedEventArgs e) => Workspace.SettingsOpen = false;
    private void SettingsClicked(object? sender, RoutedEventArgs e) => Workspace.SettingsOpen = true;
    private void ToggleThemeClicked(object? sender, RoutedEventArgs e) => Workspace.ThemeIndex = 1 - Workspace.ThemeIndex;
    private void CancelClicked(object? sender, RoutedEventArgs e) => operation?.Cancel();
    private async void ImportClicked(object? sender, RoutedEventArgs e)
    {
        if (Workspace.IsBusy) return;
        try
        {
            var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                AllowMultiple = true, Title = Workspace.Text["Import"],
                FileTypeFilter = new[] { new FilePickerFileType(Workspace.Text["Import"]) { Patterns = new[] { "*.png", "*.jpg", "*.jpeg", "*.bmp" } } }
            });
            if (closed || files.Count == 0) return;
            var paths = files.Select(file => file.TryGetLocalPath() ?? throw new IOException(Workspace.Text["LocalOnly"])).ToArray();
            Workspace.SettingsOpen = false;
            await ImportPathsAsync(paths);
        }
        catch (Exception error) { Workspace.Message = Workspace.Text["ImportErrors"] + error.Message; }
    }

    internal async Task ImportPathsAsync(IReadOnlyList<string> paths, CancellationToken cancellationToken = default)
    {
        if (Workspace.IsBusy || closed || paths.Count == 0) return;
        operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken); var token = operation.Token;
        Workspace.IsBusy = true; Workspace.Progress = 0;
        var errors = new List<string>(); var count = 0;
        try
        {
            foreach (var path in paths)
            {
                token.ThrowIfCancellationRequested();
                ImageItem? loaded = null;
                try
                {
                    loaded = await Task.Run(() => LoadImage(path), token);
                    token.ThrowIfCancellationRequested();
                    Workspace.AddImage(loaded); loaded = null;
                }
                catch (Exception error) when (error is not OperationCanceledException) { errors.Add(Path.GetFileName(path) + ": " + error.Message); }
                finally { loaded?.Dispose(); }
                Workspace.Progress = (double)++count / paths.Count;
                Workspace.Message = Workspace.Text["Importing"] + $" {count}/{paths.Count}";
            }
            Workspace.Message = errors.Count == 0 ? Workspace.Text["Imported"] : Workspace.Text["ImportErrors"] + string.Join("; ", errors);
        }
        catch (OperationCanceledException) { Workspace.Message = Workspace.Text["ImportCancelled"]; }
        finally { operation.Dispose(); operation = null; Workspace.IsBusy = false; }
    }

    private ImageItem LoadImage(string path)
    {
        ImageMetadata metadata;
        using (var bitmap = new Bitmap(path))
            metadata = new ImageMetadata(bitmap.PixelSize.Width, bitmap.PixelSize.Height, bitmap.Dpi.X, bitmap.Dpi.Y);
        using var stream = File.OpenRead(path);
        var targetWidth = Math.Max(1, (int)Math.Min(metadata.PixelWidth, Math.Min(480, 300d * metadata.PixelWidth / metadata.PixelHeight)));
        var thumbnail = Bitmap.DecodeToWidth(stream, targetWidth);
        return new ImageItem(new ImportedImage(Path.GetFullPath(path), metadata), thumbnail, Workspace.Text);
    }

    private async void ExportClicked(object? sender, RoutedEventArgs e) => await ExportAsync();
    internal async Task ExportAsync()
    {
        if (Workspace.IsBusy) return;
        var snapshot = Workspace.ExportSelection();
        if (snapshot.Length == 0) { Workspace.Message = Workspace.Text["NoSelection"]; return; }
        try
        {
            var options = Workspace.Settings.CreateOptions(Workspace.Text);
            operation = new CancellationTokenSource(); var token = operation.Token;
            Workspace.IsBusy = true; Workspace.Progress = 0; Workspace.Message = Workspace.Text["PreparingExcel"];
            var reporting = new Progress<int>(n =>
            {
                Workspace.Progress = (double)n / snapshot.Length;
                Workspace.Message = Workspace.Text["Exporting"] + $" {n}/{snapshot.Length}";
            });
            await excel.Run(() => exporter.Export(snapshot, options, excel, reporting, token));
            Workspace.Message = Workspace.Text["Exported"];
        }
        catch (OperationCanceledException) { Workspace.Message = Workspace.Text["ExportCancelled"]; }
        catch (Exception error) { Workspace.Message = Workspace.Text["ExportError"] + error.Message; }
        finally { operation?.Dispose(); operation = null; Workspace.IsBusy = false; }
    }

    private void RemoveAllClicked(object? sender, RoutedEventArgs e) { previewGeneration++; Workspace.RemoveAll(); }

    private void RemoveImageClicked(object? sender, RoutedEventArgs e)
    {
        if ((sender as Control)?.DataContext is ImageItem item) Workspace.Remove(item);
    }
    private async void CopyImageClicked(object? sender, RoutedEventArgs e)
    {
        if ((sender as Control)?.DataContext is ImageItem item) await CopyImageAsync(item);
    }
    internal async Task CopyImageAsync(ImageItem item)
    {
        Bitmap? bitmap = null;
        try
        {
            if (Clipboard is not { } clipboard) throw new InvalidOperationException(Workspace.Text["ClipboardUnavailable"]);
            bitmap = await Task.Run(() => new Bitmap(item.Image.Path));
            if (closed) return;
            // Copy source pixels, never the reduced thumbnail. Keep it alive until replacement/close.
            await clipboard.SetBitmapAsync(bitmap);
            await clipboard.FlushAsync();
            if (closed) return;
            var previous = clipboardBitmap; clipboardBitmap = bitmap; bitmap = null; previous?.Dispose();
            Workspace.Message = Workspace.Text["Copied"];
        }
        catch (Exception error) { Workspace.Message = Workspace.Text["CopyError"] + error.Message; }
        finally { bitmap?.Dispose(); }
    }

    private async void ImagePreviewClicked(object? sender, RoutedEventArgs e)
    {
        if ((sender as Control)?.DataContext is not ImageItem item) return;
        previewInvoker = sender as Control;
        await ShowPreviewAsync(item);
        ClosePreviewButton.Focus();
    }
    internal async Task ShowPreviewAsync(ImageItem item)
    {
        var generation = ++previewGeneration;
        Workspace.ClosePreview(); Workspace.PreviewItem = item; Workspace.PreviewOpen = true;
        Workspace.PreviewLoading = true; Workspace.PreviewZoom = 1;
        try
        {
            var bitmap = await Task.Run(() =>
            {
                var metadata = item.Image.Metadata;
                var targetWidth = Math.Max(1, (int)Math.Min(metadata.PixelWidth, Math.Min(2560, 2560d * metadata.PixelWidth / metadata.PixelHeight)));
                using var stream = File.OpenRead(item.Image.Path);
                return Bitmap.DecodeToWidth(stream, targetWidth);
            });
            if (closed || generation != previewGeneration || !Workspace.PreviewOpen) { bitmap.Dispose(); return; }
            Workspace.PreviewBitmap = bitmap;
        }
        catch (Exception error) { if (generation == previewGeneration) { Workspace.Message = Workspace.Text["PreviewError"] + error.Message; Workspace.ClosePreview(); } }
        finally { if (generation == previewGeneration) Workspace.PreviewLoading = false; }
    }
    private void ClosePreview()
    {
        previewGeneration++; Workspace.ClosePreview(); previewInvoker?.Focus();
    }
    private void ClosePreviewClicked(object? sender, RoutedEventArgs e) => ClosePreview();
    private void PreviewBackdropPressed(object? sender, PointerPressedEventArgs e) { if (e.Source == PreviewOverlay) ClosePreview(); }
    private void FitPreviewClicked(object? sender, RoutedEventArgs e) => Workspace.PreviewZoom = 1;
    private async void PreviousPreviewClicked(object? sender, RoutedEventArgs e) => await NavigatePreviewAsync(-1);
    private async void NextPreviewClicked(object? sender, RoutedEventArgs e) => await NavigatePreviewAsync(1);
    internal async Task NavigatePreviewAsync(int step)
    {
        if (Workspace.PreviewItem is not { } item) return;
        var index = Workspace.Images.IndexOf(item) + step;
        if (index >= 0 && index < Workspace.Images.Count) await ShowPreviewAsync(Workspace.Images[index]);
    }
    private async void WindowKeyDown(object? sender, KeyEventArgs e)
    {
        if (!Workspace.PreviewOpen) return;
        if (e.Key == Key.Escape) { e.Handled = true; ClosePreview(); }
        else if (e.Key is Key.Left or Key.Right && e.Source is not Slider)
        { e.Handled = true; await NavigatePreviewAsync(e.Key == Key.Left ? -1 : 1); }
    }
}
