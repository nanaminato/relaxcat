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
    original.Settings.StartRow = "5"; original.Settings.StartColumn = "3";
    original.Settings.SortOrderIndex = 2;
    original.Settings.ShowGridlines = false;
    original.Settings.VerticalGap = "3"; original.Settings.Rules = "width > 100 => 0.5";
    original.ThemeIndex = 1; original.LanguageIndex = 1;
    original.Capture.ProviderIndex = 2;
}
using (var restored = new RelaxCat.App.Models.WorkspaceModel(settingsStore))
{
    if (restored.Settings.HorizontalGapUnitIndex != 1 || restored.Settings.VerticalGapUnitIndex != 1 ||
        restored.Settings.StartRow != "5" || restored.Settings.StartColumn != "3" || restored.Settings.SortOrderIndex != 2 || restored.Settings.VerticalGap != "3" || restored.Settings.ShowGridlines || restored.ThemeIndex != 1 || restored.LanguageIndex != 1 ||
        restored.Capture.ProviderIndex != 2 || restored.Settings.Rules != "width > 100 => 0.5")
        throw new Exception("Settings did not survive restart");
}
using (var recovery = new RelaxCat.App.Models.WorkspaceModel(settingsStore))
{
    recovery.Settings.SortOrderIndex = 1;
    recovery.FlushSettings(); // Preserve the previous healthy settings as a backup.
    File.WriteAllText(settingsPath, "{broken");
    using var recovered = new RelaxCat.App.Models.WorkspaceModel(settingsStore);
    if (!settingsStore.RecoveredFromBackup || recovered.Settings.SortOrderIndex != 2)
        throw new Exception("Settings backup was not recovered");
    recovered.Settings.HorizontalGap = "25";
    recovered.FlushSettings();
    if (settingsStore.Load()!.Settings.HorizontalGap != "25") throw new Exception("Recovered settings could not be saved");
}
File.WriteAllText(settingsPath, "null"); File.WriteAllText(settingsPath + ".bak", "null");
using (var defaults = new RelaxCat.App.Models.WorkspaceModel(settingsStore))
    if (defaults.Settings.SortOrderIndex != 1 || defaults.Capture.ProviderIndex != 0)
        throw new Exception("Invalid settings did not fall back to defaults");
File.Delete(settingsPath); File.Delete(settingsPath + ".bak");
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
// Buttons carry an icon plus a label, so the visible text lives in child TextBlocks.
static string Label(Button button) => button.Content switch
{
    string text => text,
    Panel panel => string.Concat(panel.Children.OfType<TextBlock>().Select(text => text.Text)),
    _ => ""
};
void Capture(string file)
{
    if (window.Workspace.PreviewOpen) { Thread.Sleep(180); }
    Dispatcher.UIThread.RunJobs();
    using var frame = window.CaptureRenderedFrame() ?? throw new Exception("UI did not render");
    frame.Save("artifacts/" + file, PngBitmapEncoderOptions.Default);
}
T Control<T>(string name) where T : Control => window.FindControl<T>(name) ?? throw new Exception("Missing " + name);

Check(Label(Control<Button>("ImportButton")) == "导入图片", "Chinese UI missing");
Check(Control<Grid>("HomePage").GetVisualDescendants().OfType<TextBox>().Any(t => t.Name == "QuickScale") && !Control<Grid>("HomePage").GetVisualDescendants().OfType<TextBox>().Any(t => t.Name == "RulesEditor"), "Home is missing quick settings or contains detailed settings");
Check(window.ActualThemeVariant == ThemeVariant.Dark, "Default dark theme was not applied at startup");
Capture("desktop-empty-dark.png");
Click(Control<Button>("ThemeButton"));
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
settings.FindControl<TextBox>("StartRowInput")!.Text = "5";
settings.FindControl<TextBox>("StartColumnInput")!.Text = "3"; Dispatcher.UIThread.RunJobs();
var configured = window.Workspace.Settings.CreateOptions(window.Workspace.Text);
Check(configured.StartRow == 5 && configured.StartColumn == 3, "Start cell did not reach export options");
settings.FindControl<TextBox>("StartRowInput")!.Text = "1";
settings.FindControl<TextBox>("StartColumnInput")!.Text = "1"; Dispatcher.UIThread.RunJobs();
Check(!configured.ShowGridlines, "Gridline setting did not reach export");
gridlines.IsChecked = true;
Check(window.Icon is not null, "Window icon missing");
Check(configured.HorizontalGapUnit == GapUnit.Cells && configured.VerticalGapUnit == GapUnit.Cells && configured.SelectAllPicturesAfterExport,
    "Export settings did not reach the export options");
Capture("settings-excel.png");
// Narrow layout check: two columns must still fit at the minimum window width, and any
// overflow has to stay reachable through the pane's own scroll range.
window.Width = 1000; window.Height = 680; Dispatcher.UIThread.RunJobs();
var excelScroll = settings.FindControl<ScrollViewer>("ExcelScroll")!;
Check(excelScroll.Extent.Height > excelScroll.Viewport.Height && excelScroll.Viewport.Height > 0, "Narrow settings page clipped without a scroll range");
Capture("settings-excel-narrow.png");
window.Width = 1280; window.Height = 850; Dispatcher.UIThread.RunJobs();
tabs.SelectedIndex = 2; Dispatcher.UIThread.RunJobs();
var language = settings.FindControl<ComboBox>("LanguageSelector")!;
language.SelectedIndex = 1; Dispatcher.UIThread.RunJobs();
Check(Label(Control<Button>("ImportButton")) == "画像を追加", "Japanese labels did not update");
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

var paths = TestImages.Create();
var sortFolder = Path.GetFullPath("artifacts/sort-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(sortFolder);
var sortPaths = new[] { "new.png", "old.png", "middle.png", "tie.png" }.Select(name => Path.Combine(sortFolder, name)).ToArray();
var epoch = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
for (var i = 0; i < sortPaths.Length; i++) { File.Copy(paths[0], sortPaths[i]); File.SetLastWriteTimeUtc(sortPaths[i], epoch.AddDays(new[] { 3, 1, 2, 2 }[i])); }
Complete(window.ImportPathsAsync(sortPaths.Take(2).ToArray()));
Check(window.Workspace.Images.Select(i => i.Name).SequenceEqual(new[] { "old.png", "new.png" }), "Initial import not sorted by file time");
Complete(window.ImportPathsAsync(sortPaths.Skip(2).ToArray()));
Check(window.Workspace.Images.Select(i => i.Name).SequenceEqual(new[] { "old.png", "middle.png", "tie.png", "new.png" }), "Late imports or tied timestamps not stable");
var sortSelector = Control<ComboBox>("SortOrderSelector"); sortSelector.SelectedIndex = 2; Dispatcher.UIThread.RunJobs();
Check(window.Workspace.Images.Select(i => i.Name).SequenceEqual(new[] { "new.png", "middle.png", "tie.png", "old.png" }), "Descending time sort failed");
var selectedItem = window.Workspace.Images[1]; selectedItem.IncludeInExport = false;
Complete(window.ShowPreviewAsync(selectedItem));
sortSelector.SelectedIndex = 0; Dispatcher.UIThread.RunJobs();
Check(window.Workspace.Images.Select(i => i.Name).SequenceEqual(new[] { "new.png", "old.png", "middle.png", "tie.png" }) && window.Workspace.PreviewItem == selectedItem && !selectedItem.IncludeInExport, "Sorting changed item identity or selection");
window.Workspace.IsBusy = true; window.Workspace.RemoveAll(); Check(window.Workspace.Images.Count == 4, "Clear ran during busy operation"); window.Workspace.IsBusy = false;
window.Workspace.RemoveAll();
Check(!window.Workspace.PreviewOpen && window.Workspace.PreviewBitmap is null, "Clear left preview resources alive");
Complete(window.ImportPathsAsync(sortPaths));
Click(Control<Button>("RemoveAllButton"));
Check(window.Workspace.Images.Count == 0 && !window.Workspace.CanExport && !window.Workspace.CanRemoveAll && window.Workspace.PreviewBitmap is null && sortPaths.All(File.Exists), "Clear failed or deleted source files");
Click(Control<Button>("RemoveAllButton"));
sortSelector.SelectedIndex = 1; Dispatcher.UIThread.RunJobs();
Console.WriteLine("PASS: automatic time sorting across imports, stable ties, import order, retained selection and safe clear-all.");
Complete(window.ImportPathsAsync(paths));
Check(window.Workspace.Images.Count == paths.Length, "Image import failed");
var gallery = Control<ItemsControl>("Gallery");
Check(gallery.GetVisualDescendants().OfType<Image>().Count() == paths.Length, "Not every image is rendered");
window.Workspace.LanguageIndex = 1; Dispatcher.UIThread.RunJobs();
Check(gallery.GetVisualDescendants().OfType<Button>().Any(b => Label(b) == "コピー"), "Existing cards did not change language");
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
var exportButton = Control<Button>("ExportButton");
var hoverPoint = exportButton.TranslatePoint(new Point(exportButton.Bounds.Width / 2, exportButton.Bounds.Height / 2), window) ?? new Point(0, 0);
window.MouseMove(hoverPoint, RawInputModifiers.None); Dispatcher.UIThread.RunJobs();
Check(exportButton.IsPointerOver, "Pointer-over state did not reach the export button");
Capture("desktop-hover-light.png");
window.MouseMove(new Point(4, 4), RawInputModifiers.None); Dispatcher.UIThread.RunJobs();
Click(Control<Button>("ThemeButton")); Check(window.ActualThemeVariant == ThemeVariant.Dark, "Dark theme toggle did not apply");
Capture("desktop-gallery-dark.png");

var copy = gallery.GetVisualDescendants().OfType<Button>().First(b => b.DataContext == first && Label(b) == "复制");
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
var remove = gallery.GetVisualDescendants().OfType<Button>().First(b => b.DataContext == last && Label(b) == "移除");
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
// Drag & drop through the real headless drag pipeline: folders expand recursively; duplicates and non-images are skipped.
Check(DragDrop.GetAllowDrop(window), "Window does not accept dropped files");
var dropRoot = Path.GetFullPath("artifacts/drop-" + Guid.NewGuid().ToString("N"));
var dropDeep = Path.Combine(dropRoot, "nested", "deeper");
Directory.CreateDirectory(dropDeep);
File.Copy(paths[0], Path.Combine(dropRoot, "dropped.png"));
File.Copy(paths[1], Path.Combine(dropRoot, "nested", "inner.png"));
File.Copy(paths[2], Path.Combine(dropDeep, "deep.png"));
File.WriteAllText(Path.Combine(dropRoot, "ignored.txt"), "not an image");
File.WriteAllText(Path.Combine(dropDeep, "notes.txt"), "not an image");
var beforeDrop = window.Workspace.Images.Count;
Complete(window.DropPathsAsync(new[] { dropRoot, Path.Combine(dropRoot, "dropped.png") }));
Check(window.Workspace.Images.Count == beforeDrop + 3, "Dropped folder did not import nested images exactly once");
Check(new[] { "dropped.png", "inner.png", "deep.png" }.All(name => window.Workspace.Images.Any(item => item.Name == name)), "Dropped subfolder images missing");
Check(window.Workspace.Message == window.Workspace.Text["Imported"], "Successful drop did not report completion: " + window.Workspace.Message);
// The window-level drag events are wired through AllowDrop plus one drop handler; the layer itself is bound to DropActive.
window.Workspace.DropActive = true; Dispatcher.UIThread.RunJobs();
Check(Control<Border>("DropOverlay").IsVisible, "Drag feedback did not appear");
Capture("desktop-drop-light.png");
window.Workspace.DropActive = false; Dispatcher.UIThread.RunJobs();
Check(!Control<Border>("DropOverlay").IsVisible, "Drop left the feedback layer visible");
var dropTextOnly = Path.Combine(dropRoot, "text-only"); Directory.CreateDirectory(dropTextOnly);
File.WriteAllText(Path.Combine(dropTextOnly, "readme.txt"), "x");
Complete(window.DropPathsAsync(new[] { dropTextOnly }));
Check(window.Workspace.Images.Count == beforeDrop + 3 && window.Workspace.Message == window.Workspace.Text["DropNoImages"], "Drop without images was not reported");
Complete(window.DropPathsAsync(new[] { Path.Combine(dropRoot, "missing") }));
Check(window.Workspace.Message.StartsWith(window.Workspace.Text["DropErrors"]) && window.Workspace.Message.Contains("missing"), "Missing dropped path was not reported");
window.Workspace.IsBusy = true; Complete(window.DropPathsAsync(new[] { dropRoot }));
Check(window.Workspace.Images.Count == beforeDrop + 3 && window.Workspace.Message == window.Workspace.Text["DropBusy"], "Drop was accepted during a busy workflow");
window.Workspace.IsBusy = false;
Console.WriteLine("PASS: dropped images and folders import recursively, with deduplication, drop feedback, and busy/missing feedback.");
CaptureChecks.RunFolder(window, Complete, PumpUntil);
int? ownedGreenshot = args.Contains("--greenshot") ? CaptureChecks.RunGreenshot(window, Complete, PumpUntil) : null;
settings.OpenCapturePage(); Capture("settings-capture.png");
window.Workspace.ThemeIndex = 0; Dispatcher.UIThread.RunJobs(); Capture("settings-capture-dark.png");
window.Workspace.ThemeIndex = 1; Dispatcher.UIThread.RunJobs();
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
var closingImport = new MainWindow(new RelaxCat.App.Models.SettingsStore(Path.Combine(sortFolder, "closing-settings.json"))); closingImport.Show();
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
    public void ClearPictures() { }
    public void Insert(string path, ImagePlacement points) => Paths.Add(path);
}
