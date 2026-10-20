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
    public CaptureSettings Capture { get; } = new();
    public ObservableCollection<ImageItem> Images { get; } = new();
    public WorkspaceModel()
    {
        ThemeChoices = new[] { new LocalizedOption(Text, "Dark"), new LocalizedOption(Text, "Light") };
        GapUnitChoices = new[] { new LocalizedOption(Text, "Points"), new LocalizedOption(Text, "Cells") };
        CaptureProviderChoices = new[] { new LocalizedOption(Text, "CaptureOff"), new LocalizedOption(Text, "CaptureFolderOnly"), new LocalizedOption(Text, "CaptureGreenshot") };
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
    public bool IsBusy { get => busy; set { if (Set(ref busy, value)) { Notify(nameof(IsIdle)); Notify(nameof(CanExport)); Notify(nameof(CanApplyCapture)); } } }
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
            Notify(nameof(CaptureStatus));
            Message = Text["Ready"];
        }
    }
    public string[] Languages { get; } = { "简体中文", "日本語" };
    public LocalizedOption[] ThemeChoices { get; }
    public LocalizedOption[] GapUnitChoices { get; }
    public LocalizedOption[] CaptureProviderChoices { get; }
    private bool captureBusy, captureRunning;
    private string captureStatusKey = "CaptureStopped";
    public bool CaptureBusy { get => captureBusy; set { if (Set(ref captureBusy, value)) { Notify(nameof(CanApplyCapture)); Notify(nameof(CanStopCapture)); } } }
    public bool CaptureRunning { get => captureRunning; set { if (Set(ref captureRunning, value)) Notify(nameof(CanStopCapture)); } }
    public bool CanApplyCapture => IsIdle && !CaptureBusy;
    public bool CanStopCapture => CaptureRunning || CaptureBusy;
    public string CaptureStatusKey { get => captureStatusKey; set { if (Set(ref captureStatusKey, value)) Notify(nameof(CaptureStatus)); } }
    public string CaptureStatus => Text[CaptureStatusKey];
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
