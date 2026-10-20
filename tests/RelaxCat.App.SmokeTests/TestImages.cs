using System.Globalization;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;

internal static class TestImages
{
    public static string[] Create()
    {
        var paths = new List<string>();
        for (var i = 0; i < 4; i++)
        {
            var width = i == 3 ? 760 : 1280;
            var height = i == 3 ? 1100 : 760;
            using var bitmap = new RenderTargetBitmap(new PixelSize(width, height), new Vector(96, 96));
            using (var context = bitmap.CreateDrawingContext())
            {
                var dark = i % 2 == 0;
                IBrush Brush(string color) => new SolidColorBrush(Color.Parse(color));
                void Label(string text, double size, double x, double y, string color) => context.DrawText(
                    new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), size, Brush(color)), new Point(x, y));
                context.DrawRectangle(Brush(dark ? "#17243C" : "#F2F1EC"), null, new Rect(0, 0, width, height));
                context.DrawRectangle(Brush(dark ? "#223350" : "#FFFFFF"), null, new Rect(32, 32, width - 64, 60), 12, 12);
                Label("RELAXCAT / PROJECT NOTES", 20, 58, 48, dark ? "#B8C8E7" : "#5A6170");
                Label(new[] { "A calmer workspace.", "Make room for ideas.", "Less noise, more focus.", "Plan your next step." }[i], 42, 58, 132, dark ? "#FFFFFF" : "#263448");
                Label("Collect  /  Arrange  /  Export", 20, 60, 200, dark ? "#92ACD1" : "#778192");
                var availableWidth = width - 116;
                for (var column = 0; column < 3; column++)
                {
                    var cardWidth = (availableWidth - 32) / 3;
                    var x = 58 + column * (cardWidth + 16);
                    context.DrawRectangle(Brush(dark ? "#243755" : "#FFFFFF"), null, new Rect(x, 290, cardWidth, 220), 14, 14);
                    context.DrawRectangle(Brush(new[] { "#7297E9", "#DDA577", "#85B5A4" }[column]), null, new Rect(x + 20, 314, cardWidth - 40, 80), 8, 8);
                    Label($"0{column + 1} / Working notes", 14, x + 20, 418, dark ? "#DCE6FA" : "#354158");
                    for (var line = 0; line < 2; line++) context.DrawRectangle(Brush(dark ? "#3B506F" : "#E4E7ED"), null, new Rect(x + 20, 456 + 18 * line, cardWidth - 60 - line * 20, 6), 3, 3);
                }
                Label("Everything in its place.", 20, 58, height - 90, dark ? "#92ACD1" : "#778192");
            }
            var path = Path.GetFullPath($"artifacts/sample-{i + 1}.png");
            bitmap.Save(path, PngBitmapEncoderOptions.Default); paths.Add(path);
        }
        return paths.ToArray();
    }
}
