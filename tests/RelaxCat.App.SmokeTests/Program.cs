using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RelaxCat.App;

AppBuilder.Configure<DesktopApp>().UseHeadless(new AvaloniaHeadlessPlatformOptions
    { UseHeadlessDrawing = false }).UseSkia().SetupWithoutStarting();
var window = new MainWindow();
window.Show();
Dispatcher.UIThread.RunJobs();
if (!window.GetVisualDescendants().OfType<Button>().Any(b => Equals(b.Content, "导入图片")))
    throw new Exception("Chinese UI missing");
var language = window.GetVisualDescendants().OfType<ComboBox>().Single();
language.SelectedIndex = 1;
Dispatcher.UIThread.RunJobs();
if (!window.GetVisualDescendants().OfType<Button>().Any(b => Equals(b.Content, "画像を追加")))
    throw new Exception("Japanese UI missing");
language.SelectedIndex = 0;
Dispatcher.UIThread.RunJobs();
Directory.CreateDirectory("artifacts");
using (var frame = window.CaptureRenderedFrame())
{
    if (frame is null) throw new Exception("UI did not render");
    frame.Save("artifacts/desktop-smoke.png", Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
}
window.Close();
Console.WriteLine("PASS: desktop startup, Chinese/Japanese switching, rendered frame, clean close.");
Console.WriteLine(Type.GetTypeFromProgID("Excel.Application") is null
    ? "Excel COM unavailable on this machine; live export not verified."
    : "Excel COM registered; live validation requires separate Office integration run.");
if (args.Contains("--excel")) ExcelChecks.Run().GetAwaiter().GetResult();
