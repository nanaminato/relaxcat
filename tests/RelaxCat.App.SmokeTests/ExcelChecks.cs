using System.Reflection;
using System.Runtime.InteropServices;
using RelaxCat.App;
using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;

internal static class ExcelChecks
{
    private static dynamic Get(ExcelSession session, string field) =>
        typeof(ExcelSession).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(session)!;
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }

    public static async Task Run()
    {
        using var session = new ExcelSession();
        var path = Path.GetFullPath("artifacts/desktop-smoke.png");
        using var bitmap = new Avalonia.Media.Imaging.Bitmap(path);
        var image = new ImportedImage(path, new ImageMetadata(bitmap.PixelSize.Width, bitmap.PixelSize.Height, bitmap.Dpi.X, bitmap.Dpi.Y));
        var options = new ExportOptions(new SizingOptions(1920, 1080, .5, 1.5));
        try
        {
            await session.Run(() =>
            {
                var exporter = new ImageExporter();
                exporter.Export(new[] { image }, options, session, null, default);
                dynamic book = Get(session, "workbook");
                dynamic sheet = Get(session, "sheet");
                dynamic app = Get(session, "app");
                Console.WriteLine("Excel version: " + (string)app.Version);
                app.Visible = false;
                Console.WriteLine("CHECK: hidden test workbook");
                object? shapes = null, shape = null, rows = null, columns = null;
                try
                {
                    shapes = sheet.Shapes; shape = ((dynamic)shapes).Item(1);
                    Console.WriteLine("CHECK: shape obtained");
                    Check((int)((dynamic)shape).Placement == 1, "xlMoveAndSize missing");
                    var size = ImageSizing.Calculate(image.Metadata, options.Sizing);
                    Check(Math.Abs((double)((dynamic)shape).Width - size.DisplayWidth * .75) < .1, "Point conversion mismatch");
                    rows = sheet.Rows; columns = sheet.Columns;
                    ((dynamic)rows).RowHeight = 30; ((dynamic)columns).ColumnWidth = 15;
                    Console.WriteLine("CHECK: changed normalization settings");
                    exporter.Export(new[] { image }, options, session, null, default);
                    Console.WriteLine("CHECK: second export completed");
                    Check(ReferenceEquals((object)book, (object)Get(session, "workbook")), "Workbook was not reused");
                    Check((int)((dynamic)shapes).Count == 2, "Append failed");
                    Check(Math.Abs((double)((dynamic)rows).RowHeight - 18) < .1, "Existing workbook not normalized");
                    Check(Math.Abs((double)((dynamic)columns).ColumnWidth - 10) < .1, "Column normalization missing");
                    var template = Path.GetFullPath("artifacts/smoke-template.xltx");
                    app.DisplayAlerts = false;
                    book.SaveAs(template, 54); // xlOpenXMLTemplate
                    Console.WriteLine("CHECK: template saved");
                    book.Close(false);
                    Console.WriteLine("CHECK: workbook closed");
                    exporter.Export(new[] { image }, options with { TemplatePath = template }, session, null, default);
                    dynamic recovered = Get(session, "workbook");
                    object? recoveredSheets = null;
                    try
                    {
                        recoveredSheets = recovered.Worksheets;
                        Check((int)((dynamic)recoveredSheets).Count > 0, "Template recovery failed");
                    }
                    finally { Release(recoveredSheets); }
                    Console.WriteLine("PASS: live COM insertion, point sizing, xlMoveAndSize, reuse, normalization, closed-document recovery from template.");
                }
                finally { Release(columns); Release(rows); Release(shape); Release(shapes); }
            }).ConfigureAwait(false);
        }
        finally
        {
            await session.Run(() =>
            {
                dynamic? app = Get(session, "app");
                if (app is not null) { app.DisplayAlerts = false; app.Quit(); }
            }).ConfigureAwait(false);
        }
    }
}
