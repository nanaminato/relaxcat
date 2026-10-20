using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RelaxCat.App;
using RelaxCat.App.Views;
using RelaxCat.Core.Export;
using RelaxCat.Core.Layout;

AppBuilder.Configure<DesktopApp>().UseHeadless(new AvaloniaHeadlessPlatformOptions
    { UseHeadlessDrawing = false }).UseSkia().SetupWithoutStarting();
Directory.CreateDirectory("artifacts");
var settingsPath = Path.GetFullPath("artifacts/settings-" + Guid.NewGuid().ToString("N") + ".json");
var settingsStore = new RelaxCat.App.Models.SettingsStore(settingsPath);
using (var original = new RelaxCat.App.Models.WorkspaceModel(settingsStore))
{
    original.Settings.HorizontalGapUnitIndex = 1; original.Settings.VerticalGapUnitIndex = 1;
    original.Settings.ShowGridlines = false;
    original.Settings.VerticalGap = "3"; original.Settings.Rules = "width > 100 => 0.5";
    original.ThemeIndex = 1; original.LanguageIndex = 1;
    original.Capture.ProviderIndex = 2; original.Capture.MoveMouse = true;
}
using (var restored = new RelaxCat.App.Models.WorkspaceModel(settingsStore))
{
    if (restored.Settings.HorizontalGapUnitIndex != 1 || restored.Settings.VerticalGapUnitIndex != 1 ||
        restored.Settings.VerticalGap != "3" || restored.Settings.ShowGridlines || restored.ThemeIndex != 1 || restored.LanguageIndex != 1 ||
        restored.Capture.ProviderIndex != 2 || !restored.Capture.MoveMouse || restored.Settings.Rules != "width > 100 => 0.5")
        throw new Exception("Settings did not survive restart");
}
File.Delete(settingsPath);
Console.WriteLine("PASS: settings persistence across workspace restart.");
if (args.Contains("--settings-only")) return;
var window = new MainWindow(settingsStore); window.Show(); Dispatcher.UIThread.RunJobs();
static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
static void PumpUntil(Func<bool> done)
{
    var deadline = DateTime.UtcNow.AddSeconds(40);
    while (!done())
    {
        if (DateTime.UtcNow > deadline) throw new TimeoutException("Desktop operation timed out");
        Dispatcher.UIThread.RunJobs(); CaptureChecks.PumpNativeMessages(); Thread.Sleep(5);
    }
    Dispatcher.UIThread.RunJobs();
}
static void Complete(Task task) { PumpUntil(() => task.IsCompleted); task.GetAwaiter().GetResult(); }
static void Click(Button button) { button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Dispatcher.UIThread.RunJobs(); }
void Capture(string file)
{
    if (window.Workspace.PreviewOpen) { Thread.Sleep(180); }
    Dispatcher.UIThread.RunJobs();
    using var frame = window.CaptureRenderedFrame() ?? throw new Exception("UI did not render");
    frame.Save("artifacts/" + file, PngBitmapEncoderOptions.Default);
}
T Control<T>(string name) where T : Control => window.FindControl<T>(name) ?? throw new Exception("Missing " + name);

Check(Equals(Control<Button>("ImportButton").Content, "导入图片"), "Chinese UI missing");
Check(Control<Grid>("HomePage").GetVisualDescendants().OfType<TextBox>().Count() == 1, "Home contains detailed settings");
Capture("desktop-empty-dark.png");
Click(Control<Button>("SettingsButton"));
var settings = Control<SettingsView>("SettingsPage");
Check(settings.IsVisible && !Control<Grid>("HomePage").IsVisible, "Settings did not open as a separate page");
var tabs = settings.GetVisualDescendants().OfType<TabControl>().Single();
Capture("settings-sizing.png");
tabs.SelectedIndex = 1; Dispatcher.UIThread.RunJobs();
var horizontalUnit = settings.FindControl<ComboBox>("HorizontalGapUnitSelector")!;
var verticalUnit = settings.FindControl<ComboBox>("VerticalGapUnitSelector")!;
horizontalUnit.SelectedIndex = 1; verticalUnit.SelectedIndex = 1;
var selectAllPictures = settings.FindControl<CheckBox>("SelectAllPicturesOption")!;
selectAllPictures.IsChecked = true; Dispatcher.UIThread.RunJobs();
var gridlines = settings.FindControl<CheckBox>("ShowGridlinesOption")!;
gridlines.IsChecked = false; Dispatcher.UIThread.RunJobs();
var configured = window.Workspace.Settings.CreateOptions(window.Workspace.Text);
Check(!configured.ShowGridlines, "Gridline setting did not reach export");
gridlines.IsChecked = true;
Check(window.Icon is not null, "Window icon missing");
Check(configured.HorizontalGapUnit == GapUnit.Cells && configured.VerticalGapUnit == GapUnit.Cells && configured.SelectAllPicturesAfterExport,
    "Export settings did not reach the export options");
Capture("settings-excel.png");
tabs.SelectedIndex = 2; Dispatcher.UIThread.RunJobs();
var language = settings.FindControl<ComboBox>("LanguageSelector")!;
language.SelectedIndex = 1; Dispatcher.UIThread.RunJobs();
Check(Equals(Control<Button>("ImportButton").Content, "画像を追加"), "Japanese labels did not update");
Check(horizontalUnit.SelectedIndex == 1 && verticalUnit.SelectedIndex == 1 && window.Workspace.Settings.SelectAllPicturesAfterExport,
    "Language change reset gap units or export selection option");
Check(window.Workspace.Text["Remove"] == "削除", "Card translations missing");
var theme = settings.FindControl<ComboBox>("ThemeSelector")!;
theme.SelectedIndex = 1; Dispatcher.UIThread.RunJobs();
Check(window.ActualThemeVariant == ThemeVariant.Light, "Light theme did not apply");
Capture("settings-appearance-light-ja.png");
language.SelectedIndex = 0; Dispatcher.UIThread.RunJobs();
Check(window.Workspace.ThemeIndex == 1 && theme.SelectedIndex == 1 && window.ActualThemeVariant == ThemeVariant.Light, "Language change reset theme");
horizontalUnit.SelectedIndex = 0; verticalUnit.SelectedIndex = 0; selectAllPictures.IsChecked = false;
Click(Control<Button>("HomeButton")); Capture("desktop-empty-light.png");

var paths = TestImages.Create(); Complete(window.ImportPathsAsync(paths));
Check(window.Workspace.Images.Count == paths.Length, "Image import failed");
var gallery = Control<ItemsControl>("Gallery");
Check(gallery.GetVisualDescendants().OfType<Image>().Count() == paths.Length, "Not every image is rendered");
window.Workspace.LanguageIndex = 1; Dispatcher.UIThread.RunJobs();
Check(gallery.GetVisualDescendants().OfType<Button>().Any(b => Equals(b.Content, "コピー")), "Existing cards did not change language");
window.Workspace.LanguageIndex = 0; Dispatcher.UIThread.RunJobs();
Control<TextBox>("QuickScale").Text = "0.75"; Dispatcher.UIThread.RunJobs();
Check(window.Workspace.Settings.CreateOptions(window.Workspace.Text).Sizing.DesignScale == .75, "Quick setting did not reach export configuration");
window.Workspace.Settings.DesignScale = "0.5"; Dispatcher.UIThread.RunJobs();
var first = window.Workspace.Images[0];
var checkBox = gallery.GetVisualDescendants().OfType<CheckBox>().First(c => c.DataContext == first);
checkBox.IsChecked = false; Dispatcher.UIThread.RunJobs();
Check(!first.IncludeInExport && window.Workspace.ExportSelection().Length == paths.Length - 1, "Export checkbox did not filter selection");
Check(Control<TextBlock>("SummaryLabel").Text!.Contains($"待导出 {paths.Length - 1}"), "Selection count did not update");
var target = new RecordingTarget();
new ImageExporter().Export(window.Workspace.ExportSelection(), window.Workspace.Settings.CreateOptions(window.Workspace.Text), target, null, default);
Check(!target.Paths.Contains(first.Image.Path) && target.Paths.Count == paths.Length - 1, "Unchecked image reached exporter");
Capture("desktop-gallery-light.png");
Click(Control<Button>("ThemeButton")); Check(window.ActualThemeVariant == ThemeVariant.Dark, "Dark theme toggle did not apply");
Capture("desktop-gallery-dark.png");

var copy = gallery.GetVisualDescendants().OfType<Button>().First(b => b.DataContext == first && Equals(b.Content, "复制"));
Click(copy); PumpUntil(() => window.Workspace.Message == window.Workspace.Text["Copied"] || window.Workspace.Message.StartsWith(window.Workspace.Text["CopyError"]));
Check(window.Workspace.Message == window.Workspace.Text["Copied"], window.Workspace.Message);
var copiedTask = window.Clipboard!.TryGetBitmapAsync(); Complete(copiedTask);
using (var copied = copiedTask.GetAwaiter().GetResult())
    Check(copied?.PixelSize.Width == first.Image.Metadata.PixelWidth && copied?.PixelSize.Height == first.Image.Metadata.PixelHeight, "Clipboard contains a thumbnail instead of source pixels");

var preview = gallery.GetVisualDescendants().OfType<Button>().First(b => b.DataContext == first && b.Content is Border);
Click(preview); PumpUntil(() => !window.Workspace.PreviewLoading);
Check(window.Workspace.PreviewOpen && Control<Image>("LargePreview").Source is not null, "Preview did not open");
Check(!window.Workspace.HasPrevious && window.Workspace.HasNext, "Preview boundaries incorrect");
Control<Slider>("PreviewZoomSlider").Value = 1.5; Dispatcher.UIThread.RunJobs();
Check(Math.Abs(window.Workspace.PreviewZoom - 1.5) < .01, "Zoom binding failed"); Capture("desktop-preview-dark.png");
Click(Control<Button>("NextPreviewButton")); PumpUntil(() => !window.Workspace.PreviewLoading);
Check(window.Workspace.PreviewItem == window.Workspace.Images[1], "Next preview failed");
var previewCheckbox = Control<Border>("PreviewOverlay").GetVisualDescendants().OfType<CheckBox>().Single();
previewCheckbox.IsChecked = false; Dispatcher.UIThread.RunJobs();
Check(!window.Workspace.Images[1].IncludeInExport && window.Workspace.ExportSelection().Length == paths.Length - 2, "Preview checkbox did not update export selection");
window.Workspace.ThemeIndex = 1; Dispatcher.UIThread.RunJobs(); Capture("desktop-preview-light.png");
window.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape }); Dispatcher.UIThread.RunJobs();
Check(!window.Workspace.PreviewOpen && window.Workspace.PreviewBitmap is null, "Esc did not close/release preview");

var last = window.Workspace.Images[^1];
var remove = gallery.GetVisualDescendants().OfType<Button>().First(b => b.DataContext == last && Equals(b.Content, "移除"));
Click(remove);
Check(window.Workspace.Images.Count == paths.Length - 1 && File.Exists(last.Image.Path), "Remove deleted source or did not update gallery");
foreach (var item in window.Workspace.Images) item.IncludeInExport = false;
Complete(window.ExportAsync());
Check(!window.Workspace.CanExport && window.Workspace.Message == window.Workspace.Text["NoSelection"], "Empty selection is exportable");
var corrupt = Path.GetFullPath("artifacts/corrupt.png"); File.WriteAllText(corrupt, "invalid image");
var previousCount = window.Workspace.Images.Count; Complete(window.ImportPathsAsync(new[] { corrupt }));
Check(window.Workspace.Images.Count == previousCount && window.Workspace.Message.StartsWith(window.Workspace.Text["ImportErrors"]), "Corrupt file handling failed");
window.Workspace.Settings.Rules = "width > 0\nscale = 1\ninvalid\nscale = 2";
CheckThrows(() => window.Workspace.Settings.CreateOptions(window.Workspace.Text)); window.Workspace.Settings.Rules = "";
CaptureChecks.RunFolder(window, Complete, PumpUntil);
int? ownedGreenshot = args.Contains("--greenshot") ? CaptureChecks.RunGreenshot(window, Complete, PumpUntil) : null;
settings.OpenCapturePage(); Capture("settings-capture.png");
window.Width = 1000; window.Height = 680; window.Workspace.LanguageIndex = 1;
Capture("settings-capture-compact-ja.png");
Check(settings.FindControl<Button>("ApplyCaptureButton")!.IsEffectivelyVisible, "Capture action disappeared at minimum window size");
window.Width = 1280; window.Height = 850; window.Workspace.LanguageIndex = 0;
Capture("desktop-smoke.png"); window.Close(); PumpUntil(() => !window.IsVisible);
if (ownedGreenshot.HasValue)
{
    try { using var process = System.Diagnostics.Process.GetProcessById(ownedGreenshot.Value); Check(process.HasExited, "Greenshot survived RelaxCat shutdown"); }
    catch (ArgumentException) { }
    Console.WriteLine("PASS: closing RelaxCat stops the owned Greenshot process.");
}
Console.WriteLine("PASS: XAML startup, separate settings pages, both themes/languages, gallery import, per-image export filtering, source-image clipboard copy, removal, zoom/navigation/Esc preview, corrupt image and rule validation.");
var closingImport = new MainWindow(); closingImport.Show();
var activeImport = closingImport.ImportPathsAsync(paths);
closingImport.Close(); PumpUntil(() => !closingImport.IsVisible);
Check(activeImport.IsCompleted && !closingImport.Workspace.IsBusy, "Closing during import did not cancel and finish cleanly");
Console.WriteLine("PASS: closing during import waits for cancellation and exits without a second close request.");
if (args.Contains("--excel")) ExcelChecks.Run().GetAwaiter().GetResult();

static void CheckThrows(Action action)
{
    try { action(); } catch (RelaxCat.Core.Imaging.ScaleRuleException) { return; }
    throw new Exception("Invalid rules were accepted");
}
internal sealed class RecordingTarget : IExcelTarget
{
    public List<string> Paths { get; } = new();
    public double NextTop => 0;
    public ExcelCellSize CellSize => new(60, 18);
    public void SelectAllPictures() { }
    public void Prepare(string? templatePath, double rowHeight, double columnWidth, bool showGridlines = true) { }
    public void Insert(string path, ImagePlacement points) => Paths.Add(path);
}
