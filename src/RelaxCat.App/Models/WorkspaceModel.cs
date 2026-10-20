using System.Collections.ObjectModel;
using System.ComponentModel;
using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using RelaxCat.Core.Export;

namespace RelaxCat.App.Models;

public sealed class WorkspaceModel : ObservableModel, IDisposable
{
    private bool settingsOpen, busy, previewOpen, previewLoading;
    private int languageIndex, themeIndex;
    private string message = "";
    private double progress, previewZoom = 1;
    private ImageItem? previewItem;
    private Bitmap? previewBitmap;
    public UiText Text { get; } = new();
    public SettingsModel Settings { get; } = new();
    public ObservableCollection<ImageItem> Images { get; } = new();
    public WorkspaceModel()
    {
        ThemeChoices = new[] { new ThemeOption(Text, "Dark"), new ThemeOption(Text, "Light") };
        Message = Text["Ready"];
        Images.CollectionChanged += (_, e) =>
        {
            if (e.OldItems is not null) foreach (ImageItem item in e.OldItems) item.PropertyChanged -= ItemChanged;
            if (e.NewItems is not null) foreach (ImageItem item in e.NewItems) item.PropertyChanged += ItemChanged;
            UpdateCounts();
        };
    }
    private void ItemChanged(object? sender, PropertyChangedEventArgs e) { if (e.PropertyName == nameof(ImageItem.IncludeInExport)) UpdateCounts(); }
    private void UpdateCounts() { Notify(nameof(Summary)); Notify(nameof(NoImages)); Notify(nameof(CanExport)); Notify(nameof(HasPrevious)); Notify(nameof(HasNext)); }
    public string Summary => string.Format(Text["Count"], Images.Count, Images.Count(i => i.IncludeInExport));
    public bool NoImages => Images.Count == 0;
    public bool CanExport => IsIdle && Images.Any(i => i.IncludeInExport);
    public bool IsBusy { get => busy; set { if (Set(ref busy, value)) { Notify(nameof(IsIdle)); Notify(nameof(CanExport)); } } }
    public bool IsIdle => !IsBusy;
    public bool SettingsOpen { get => settingsOpen; set { if (Set(ref settingsOpen, value)) Notify(nameof(HomeOpen)); } }
    public bool HomeOpen => !SettingsOpen;
    public int LanguageIndex
    {
        get => languageIndex;
        set
        {
            if (value is < 0 or > 1) return;
            if (!Set(ref languageIndex, value)) return;
            Text.Japanese = value == 1; Notify(nameof(Summary));
            Message = Text["Ready"];
        }
    }
    public string[] Languages { get; } = { "简体中文", "日本語" };
    public ThemeOption[] ThemeChoices { get; }
    public int ThemeIndex
    {
        get => themeIndex;
        set
        {
            if (value is < 0 or > 1) return;
            if (!Set(ref themeIndex, value)) return;
            if (Application.Current is { } app) app.RequestedThemeVariant = value == 1 ? ThemeVariant.Light : ThemeVariant.Dark;
        }
    }
    public string Message { get => message; set => Set(ref message, value); }
    public double Progress { get => progress; set => Set(ref progress, value); }
    public bool PreviewOpen { get => previewOpen; set => Set(ref previewOpen, value); }
    public ImageItem? PreviewItem { get => previewItem; set { if (Set(ref previewItem, value)) { Notify(nameof(HasPrevious)); Notify(nameof(HasNext)); } } }
    public Bitmap? PreviewBitmap { get => previewBitmap; set => Set(ref previewBitmap, value); }
    public bool PreviewLoading { get => previewLoading; set => Set(ref previewLoading, value); }
    public bool HasPrevious => PreviewItem is { } item && Images.IndexOf(item) > 0;
    public bool HasNext => PreviewItem is { } item && Images.IndexOf(item) < Images.Count - 1;
    public double PreviewZoom { get => previewZoom; set => Set(ref previewZoom, Math.Clamp(value, .5, 3)); }
    public ImportedImage[] ExportSelection() => Images.Where(i => i.IncludeInExport).Select(i => i.Image).ToArray();
    public void Remove(ImageItem item)
    {
        if (IsBusy || !Images.Remove(item)) return;
        if (PreviewItem == item) ClosePreview();
        item.Dispose();
    }
    public void ClosePreview()
    {
        PreviewOpen = false; PreviewItem = null;
        var old = PreviewBitmap; PreviewBitmap = null; old?.Dispose();
    }
    public void Dispose() { ClosePreview(); foreach (var image in Images) image.Dispose(); }
}
