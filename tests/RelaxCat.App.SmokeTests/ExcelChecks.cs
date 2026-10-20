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
        var options = new ExportOptions(new SizingOptions(1920, 1080, .5, 1.5), StartRow: 5, StartColumn: 3);
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
                object? excelWindow = null;
                try
                {
                    session.Prepare(null, options.RowHeight, options.ColumnWidth, false);
                    excelWindow = app.ActiveWindow;
                    Check(!(bool)((dynamic)excelWindow).DisplayGridlines, "Gridlines were not hidden");
                    // Exercise both states without adding extra images to the geometry assertions below.
                    session.Prepare(null, options.RowHeight, options.ColumnWidth, true);
                    Check((bool)((dynamic)excelWindow).DisplayGridlines, "Gridlines were not restored");
                }
                finally { Release(excelWindow); }
                app.Visible = false;
                Console.WriteLine("CHECK: hidden test workbook");
                object? shapes = null, shape = null, rows = null, columns = null;
                try
                {
                    shapes = sheet.Shapes; shape = ((dynamic)shapes).Item(1);
                    var origin = session.GetCellOrigin(5, 3);
                    Check(Math.Abs((double)((dynamic)shape).Left - origin.Left) < .1 &&
                        Math.Abs((double)((dynamic)shape).Top - origin.Top) < .1, "Start cell C5 was not used");
                    Console.WriteLine("PASS: first image starts at actual Excel cell C5.");
                    Console.WriteLine("CHECK: shape obtained");
                    Check((int)((dynamic)shape).Placement == 1, "xlMoveAndSize missing");
                    Check((int)((dynamic)shape).LockAspectRatio == -1, "Image aspect ratio is not locked");
                    var size = ImageSizing.Calculate(image.Metadata, options.Sizing);
                    Check(Math.Abs((double)((dynamic)shape).Width - size.DisplayWidth * .75) < .1, "Point conversion mismatch");
                    var originalWidth = (double)((dynamic)shape).Width;
                    var originalHeight = (double)((dynamic)shape).Height;
                    exporter.Export(new[] { image }, options with { RowHeight = 25, ColumnWidth = 2.38 }, session, null, default);
                    var changedWidth = (double)((dynamic)shape).Width;
                    var changedHeight = (double)((dynamic)shape).Height;
                    Console.WriteLine($"CHECK: normalization geometry {originalWidth:F2}x{originalHeight:F2} -> {changedWidth:F2}x{changedHeight:F2}");
                    Check(Math.Abs(originalWidth - changedWidth) < .1 && Math.Abs(originalHeight - changedHeight) < .1,
                        "Worksheet normalization distorted an existing image");
                    Check((int)((dynamic)shape).Placement == 1 && (int)((dynamic)shape).LockAspectRatio == -1,
                        "Normalization did not restore placement/aspect locking");
                    rows = sheet.Rows; columns = sheet.Columns;
                    ((dynamic)rows).RowHeight = 30; ((dynamic)columns).ColumnWidth = 15;
                    Console.WriteLine("CHECK: changed normalization settings");
                    exporter.Export(new[] { image }, options, session, null, default);
                    Console.WriteLine("CHECK: second export completed");
                    Check(ReferenceEquals((object)book, (object)Get(session, "workbook")), "Workbook was not reused");
                    Check((int)((dynamic)shapes).Count == 3, "Append failed");
                    Check(Math.Abs((double)((dynamic)rows).RowHeight - 18) < .1, "Existing workbook not normalized");
                    Check(Math.Abs((double)((dynamic)columns).ColumnWidth - 10) < .1, "Column normalization missing");
                    CheckCellGapsAndSelection(session, exporter, image, options);
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

    private static void CheckCellGapsAndSelection(ExcelSession session, ImageExporter exporter, ImportedImage image, ExportOptions options)
    {
        dynamic sheet = Get(session, "sheet"); dynamic app = Get(session, "app");
        object? shapes = null, annotation = null, first = null, second = null, selection = null, selected = null;
        try
        {
            shapes = sheet.Shapes;
            annotation = ((dynamic)shapes).AddShape(1, 0, 0, 10, 10); // Non-picture must not be selected.
            var count = (int)((dynamic)shapes).Count;
            var cells = options with
            {
                Sizing = options.Sizing with { Rules = "width > 0\nscale = 0.1" },
                RowHeight = 18, ColumnWidth = 2.38, SpaceSaving = true,
                HorizontalGap = 1.5, VerticalGap = 2,
                HorizontalGapUnit = GapUnit.Cells, VerticalGapUnit = GapUnit.Cells,
                SelectAllPicturesAfterExport = true
            };
            exporter.Export(new[] { image, image }, cells, session, null, default);
            first = ((dynamic)shapes).Item(count + 1); second = ((dynamic)shapes).Item(count + 2);
            var actualHorizontalGap = (double)((dynamic)second).Left - (double)((dynamic)first).Left - (double)((dynamic)first).Width;
            Check(Math.Abs(actualHorizontalGap - 1.5 * session.CellSize.WidthPoints) < .1, "Cell-width gap conversion failed");
            Check(Math.Abs(session.CellSize.WidthPoints - 2.38) > 1, "Character width was mistaken for points");
            selection = app.Selection; selected = ((dynamic)selection).ShapeRange;
            Check((int)((dynamic)selected).Count == count + 1, "Select all did not select exactly the pictures");
            for (var i = 1; i <= (int)((dynamic)selected).Count; i++)
            {
                object? picture = null;
                try { picture = ((dynamic)selected).Item(i); Check((int)((dynamic)picture).Type is 11 or 13, "Non-picture was selected"); }
                finally { Release(picture); }
            }
            Release(second); second = null; Release(first); first = null;
            Release(selected); selected = null; Release(selection); selection = null;
            count = (int)((dynamic)shapes).Count;
            var previousBottom = session.NextTop;
            exporter.Export(new[] { image, image }, cells with { SpaceSaving = false, SelectAllPicturesAfterExport = false }, session, null, default);
            first = ((dynamic)shapes).Item(count + 1); second = ((dynamic)shapes).Item(count + 2);
            var actualVerticalGap = (double)((dynamic)second).Top - (double)((dynamic)first).Top - (double)((dynamic)first).Height;
            Check(Math.Abs(actualVerticalGap - 2 * session.CellSize.HeightPoints) < .1, "Cell-height gap conversion failed");
            Check(Math.Abs((double)((dynamic)first).Top - previousBottom - 2 * session.CellSize.HeightPoints) < .1, "Append did not use configured vertical gap");
            Console.WriteLine($"PASS: cell gaps ({session.CellSize.WidthPoints:F2} pt/column, {session.CellSize.HeightPoints:F2} pt/row), fractional spacing, append gap, select all pictures excluding non-pictures.");
        }
        finally { Release(selected); Release(selection); Release(second); Release(first); Release(annotation); Release(shapes); }
    }
}
