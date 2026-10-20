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
        var image = new ImportedImage(path, new ImageMetadata(400, 200));
        var options = new ExportOptions(new SizingOptions(1920, 1080, .5, 1.5), StartRow: 5, StartColumn: 3);
        try
        {
            await session.Run(() =>
            {
                var exporter = new ImageExporter();
                exporter.Export(new[] { image, image }, options, session, null, default);
                dynamic app = Get(session, "app"); dynamic sheet = Get(session, "sheet"); dynamic book = Get(session, "workbook");
                object? shapes = null, note = null, cell = null, cells = null, picture = null, otherSheets = null, other = null, otherShapes = null, otherPicture = null;
                try
                {
                    app.DisplayAlerts = false;
                    shapes = sheet.Shapes;
                    Check((int)((dynamic)shapes).Count == 2, "Initial export failed");
                    note = ((dynamic)shapes).AddShape(1, 10, 10, 20, 20);
                    cells = sheet.Cells; cell = ((dynamic)cells)[1, 1]; ((dynamic)cell).Value2 = "keep";
                    otherSheets = book.Worksheets; other = ((dynamic)otherSheets).Add(Type.Missing, sheet); otherShapes = ((dynamic)other).Shapes;
                    otherPicture = ((dynamic)otherShapes).AddPicture(path, 0, -1, 0, 0, 100, 100);
                    exporter.Export(new[] { image }, options with { ShowGridlines = false, SelectAllPicturesAfterExport = true }, session, null, default);
                    Check(ReferenceEquals((object)book, (object)Get(session, "workbook")), "Workbook not reused");
                    Check((int)((dynamic)shapes).Count == 2, "Old pictures were not replaced or annotation was removed");
                    Check((string)((dynamic)cell).Value2 == "keep", "Cell content was cleared");
                    Check((int)((dynamic)otherShapes).Count == 1, "Other worksheet pictures were cleared");
                    picture = ((dynamic)shapes).Item(2);
                    var origin = session.GetCellOrigin(5, 3);
                    Check(Math.Abs((double)((dynamic)picture).Left - origin.Left) < .1 && Math.Abs((double)((dynamic)picture).Top - origin.Top) < .1, "Replacement did not restart at C5");
                    Check((int)((dynamic)picture).Placement == 1 && (int)((dynamic)picture).LockAspectRatio == -1, "Picture sizing settings missing");
                    var template = Path.GetFullPath("artifacts/smoke-template.xltx"); book.SaveAs(template, 54); book.Close(false);
                    exporter.Export(new[] { image }, options with { TemplatePath = template }, session, null, default);
                    object? recoveredShapes = null;
                    try
                    {
                        recoveredShapes = Get(session, "sheet").Shapes;
                        Check((int)((dynamic)recoveredShapes).Count == 2, "Template pictures were not replaced");
                    }
                    finally { Release(recoveredShapes); }
                    Console.WriteLine("PASS: live Excel replacement at C5, workbook reuse, preserved cells/non-picture shapes/other worksheet, picture placement and template recovery.");
                }
                finally { Release(otherPicture); Release(otherShapes); Release(other); Release(otherSheets); Release(picture); Release(cell); Release(cells); Release(note); Release(shapes); }
            }).ConfigureAwait(false);
        }
        finally { await session.Run(() => { dynamic? app = Get(session, "app"); if (app is not null) { app.DisplayAlerts = false; app.Quit(); } }).ConfigureAwait(false); }
    }
}
