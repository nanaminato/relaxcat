using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using RelaxCat.Core.Export;
using RelaxCat.Core.Layout;

namespace RelaxCat.App;

// All RCWs, including their cleanup, stay on this single STA thread.
public sealed class ExcelSession : IExcelTarget, IDisposable
{
    private readonly BlockingCollection<Action> work = new();
    private readonly Thread thread;
    private dynamic? app, workbook, sheet;
    private double nextTop;
    public double NextTop => nextTop;
    public ExcelCellSize CellSize { get; private set; }

    public ExcelSession()
    {
        thread = new Thread(() => { foreach (var action in work.GetConsumingEnumerable()) action(); })
            { IsBackground = true, Name = "RelaxCat Excel STA" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task Run(Action action)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        work.Add(() => { try { action(); completion.SetResult(); } catch (Exception e) { completion.SetException(e); } });
        return completion.Task;
    }

    public void Prepare(string? templatePath, double rowHeight, double columnWidth, bool showGridlines = true)
    {
        if (Thread.CurrentThread != thread) throw new InvalidOperationException("Excel requires its STA worker.");
        if (!double.IsFinite(rowHeight) || rowHeight <= 0 || rowHeight > 409 ||
            !double.IsFinite(columnWidth) || columnWidth <= 0 || columnWidth > 255)
            throw new ArgumentOutOfRangeException(nameof(rowHeight));
        var usable = false;
        try { if (workbook is not null && sheet is not null) { _ = workbook.Name; _ = sheet.Name; usable = true; } }
        catch (COMException) { }
        if (!usable)
        {
            if (!string.IsNullOrWhiteSpace(templatePath) && !File.Exists(templatePath))
                throw new FileNotFoundException("Excel template", templatePath);
            Release(sheet); sheet = null; Release(workbook); workbook = null;
            try { if (app is not null) _ = app.Version; }
            catch (COMException) { Release(app); app = null; }
            app ??= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel COM is unavailable."));
            object? books = null, sheets = null;
            try
            {
                books = app!.Workbooks;
                workbook = string.IsNullOrWhiteSpace(templatePath)
                    ? ((dynamic)books).Add() : ((dynamic)books).Add(Path.GetFullPath(templatePath));
                sheets = workbook.Worksheets;
                sheet = ((dynamic)sheets)[1];
                nextTop = 0;
            }
            finally { Release(sheets); Release(books); }
        }
        NormalizePreservingPictures(rowHeight, columnWidth);
        object? cells = null, cell = null;
        try
        {
            cells = sheet!.Cells; cell = ((dynamic)cells)[1, 1];
            CellSize = new ExcelCellSize((double)((dynamic)cell).Width, (double)((dynamic)cell).Height);
            app!.Visible = true;
            workbook!.Activate();
            sheet!.Activate();
            object? window = null;
            try
            {
                window = app.ActiveWindow;
                ((dynamic)window).DisplayGridlines = showGridlines;
            }
            finally { Release(window); }
        }
        finally { Release(cell); Release(cells); }
        // Normalization or user edits may have moved/resized earlier shapes.
        // Read the actual worksheet instead of trusting the last exported position.
        object? existing = null;
        try
        {
            existing = sheet!.Shapes;
            nextTop = 0;
            for (var i = 1; i <= (int)((dynamic)existing).Count; i++)
            {
                object? item = null;
                try
                {
                    item = ((dynamic)existing).Item(i);
                    nextTop = Math.Max(nextTop, (double)((dynamic)item).Top + (double)((dynamic)item).Height);
                }
                finally { Release(item); }
            }
        }
        finally { Release(existing); }
    }

    private sealed record PictureGeometry(object Shape, double Left, double Top, double Width, double Height, int Placement, int LockAspectRatio);
    private static bool IsPicture(dynamic shape) => (int)shape.Type is 11 or 13; // msoLinkedPicture / msoPicture

    private void NormalizePreservingPictures(double rowHeight, double columnWidth, bool showGridlines = true)
    {
        var saved = new List<PictureGeometry>();
        object? shapes = null, rows = null, columns = null;
        try
        {
            shapes = sheet!.Shapes;
            for (var i = 1; i <= (int)((dynamic)shapes).Count; i++)
            {
                object? shape = ((dynamic)shapes).Item(i);
                try
                {
                    if (!IsPicture(shape)) continue;
                    dynamic picture = shape;
                    saved.Add(new PictureGeometry(shape, (double)picture.Left, (double)picture.Top,
                        (double)picture.Width, (double)picture.Height, (int)picture.Placement, (int)picture.LockAspectRatio));
                    shape = null; // The snapshot owns this RCW until restoration.
                    picture.Placement = 3; // Temporarily xlFreeFloating, only during our normalization.
                }
                finally { Release(shape); }
            }
            rows = sheet.Rows; columns = sheet.Columns;
            ((dynamic)rows).RowHeight = rowHeight;
            ((dynamic)columns).ColumnWidth = columnWidth;
        }
        finally
        {
            Release(columns); Release(rows); Release(shapes);
            Exception? failure = null;
            foreach (var geometry in saved)
            {
                dynamic picture = geometry.Shape;
                try
                {
                    try
                    {
                        picture.LockAspectRatio = 0;
                        picture.Width = (float)geometry.Width; picture.Height = (float)geometry.Height;
                        picture.Left = (float)geometry.Left; picture.Top = (float)geometry.Top;
                    }
                    finally
                    {
                        try { picture.LockAspectRatio = geometry.LockAspectRatio; }
                        finally { picture.Placement = geometry.Placement; }
                    }
                }
                catch (Exception error) { failure ??= error; }
                finally { Release(geometry.Shape); }
            }
            if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }

    public void Insert(string path, ImagePlacement points)
    {
        object? shapes = null, shape = null;
        try
        {
            shapes = sheet!.Shapes;
            shape = ((dynamic)shapes).AddPicture(Path.GetFullPath(path), 0, -1,
                (float)points.X, (float)points.Y, (float)points.Width, (float)points.Height);
            ((dynamic)shape).LockAspectRatio = 0;
            ((dynamic)shape).Width = (float)points.Width; ((dynamic)shape).Height = (float)points.Height;
            ((dynamic)shape).LockAspectRatio = -1; // msoTrue, independent of cell-following placement.
            ((dynamic)shape).Placement = 1; // xlMoveAndSize, available since Excel 2007.
            nextTop = Math.Max(nextTop, points.Y + points.Height);
        }
        catch
        {
            if (shape is not null) ((dynamic)shape).Delete();
            throw;
        }
        finally { Release(shape); Release(shapes); }
    }

    public void SelectAllPictures()
    {
        object? shapes = null, range = null;
        var names = new List<object>();
        try
        {
            shapes = sheet!.Shapes;
            for (var i = 1; i <= (int)((dynamic)shapes).Count; i++)
            {
                object? shape = null;
                try { shape = ((dynamic)shapes).Item(i); if (IsPicture(shape)) names.Add((string)((dynamic)shape).Name); }
                finally { Release(shape); }
            }
            if (names.Count == 0) return;
            sheet.Activate();
            range = ((dynamic)shapes).Range(names.ToArray());
            ((dynamic)range).Select();
        }
        finally { Release(range); Release(shapes); }
    }

    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    public void Dispose()
    {
        work.Add(() =>
        {
            try
            {
                // A failed startup must not leave an invisible empty Excel process.
                if (app is not null)
                {
                    object? books = null;
                    try { books = app.Workbooks; if ((int)((dynamic)books).Count == 0) app.Quit(); }
                    finally { Release(books); }
                }
            }
            catch (COMException) { /* Excel may already have exited. */ }
            finally { Release(sheet); Release(workbook); Release(app); sheet = workbook = app = null; }
        });
        work.CompleteAdding(); // Leave visible Excel documents open for the user to save.
        thread.Join(TimeSpan.FromSeconds(2));
    }
}
