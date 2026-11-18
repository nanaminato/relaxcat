using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using RelaxCat.Core.Export;
using RelaxCat.Core.Layout;

namespace RelaxCat.App;

// All RCWs, including their cleanup, stay on this single STA thread.
public sealed class ExcelSession : IExcelTarget, IDisposable
{
    private const int ExcelApplicationError = unchecked((int)0x800A03EC);
    private const int XlWbatWorksheet = -4167;
    private readonly BlockingCollection<Action> work = new();
    private readonly Thread thread;
    private dynamic? app, workbook, sheet;
    private CancellationToken activeCancellation;
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

    public Task Run(Action action, CancellationToken cancellationToken = default)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        work.Add(() =>
        {
            activeCancellation = cancellationToken;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                action();
                cancellationToken.ThrowIfCancellationRequested();
                completion.SetResult();
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            { completion.SetCanceled(cancellationToken); }
            catch (Exception e) { completion.SetException(e); }
            finally { activeCancellation = default; }
        });
        return completion.Task;
    }

    public void Prepare(string? templatePath, double rowHeight, double columnWidth, bool showGridlines = true) =>
        Prepare(templatePath, rowHeight, columnWidth, showGridlines, createNew: false);

    internal void PrepareNewWorkbook(ExportOptions options) =>
        Prepare(options.TemplatePath, options.RowHeight, options.ColumnWidth, options.ShowGridlines, createNew: true);

    private void Prepare(string? templatePath, double rowHeight, double columnWidth, bool showGridlines, bool createNew)
    {
        if (Thread.CurrentThread != thread) throw new InvalidOperationException("Excel requires its STA worker.");
        if (!double.IsFinite(rowHeight) || rowHeight <= 0 || rowHeight > 409 ||
            !double.IsFinite(columnWidth) || columnWidth <= 0 || columnWidth > 255)
            throw new ArgumentOutOfRangeException(nameof(rowHeight));
        PrepareWorkbook(templatePath, createNew);
        activeCancellation.ThrowIfCancellationRequested();
        try { PrepareWorksheet(rowHeight, columnWidth, showGridlines); }
        catch (Exception error) when (error is not OperationCanceledException)
        { throw new ExcelExportException("ExportStageWorksheet", error); }
    }

    private void PrepareWorkbook(string? templatePath, bool createNew)
    {
        var usable = false;
        try { if (!createNew && workbook is not null && sheet is not null) { _ = workbook.Name; _ = sheet.Name; usable = true; } }
        catch (COMException) { }
        if (!usable)
        {
            if (!string.IsNullOrWhiteSpace(templatePath) && !File.Exists(templatePath))
                throw new FileNotFoundException("Excel template", templatePath);
            if (!createNew)
            {
                Release(sheet); sheet = null;
                Release(workbook); workbook = null;
            }
            try { if (app is not null) _ = app.Version; }
            catch (COMException)
            {
                Release(sheet); sheet = null;
                Release(workbook); workbook = null;
                Release(app); app = null;
            }
            try
            {
                app ??= Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Microsoft Excel COM is unavailable."));
            }
            catch (Exception error) { throw new ExcelExportException("ExportStageStartExcel", error); }
            object? books = null, sheets = null, newWorkbook = null;
            var stage = "ExportStageWorkbook";
            try
            {
                books = app!.Workbooks;
                newWorkbook = CreateWorkbook(books, templatePath);
                activeCancellation.ThrowIfCancellationRequested();
                stage = "ExportStageWorksheet";
                sheets = ((dynamic)newWorkbook).Worksheets;
                object newSheet = ((dynamic)sheets)[1];
                Release(sheet); Release(workbook);
                sheet = newSheet;
                workbook = newWorkbook; newWorkbook = null;
                nextTop = 0;
            }
            catch (Exception error)
            {
                if (activeCancellation.IsCancellationRequested && newWorkbook is not null)
                {
                    try { ((dynamic)newWorkbook).Close(false); }
                    catch (Exception) { /* Keep the cancellation result. */ }
                }
                if (workbook is null) ReleaseFailedStartup(books);
                activeCancellation.ThrowIfCancellationRequested();
                throw new ExcelExportException(stage, error);
            }
            finally { Release(sheets); Release(newWorkbook); Release(books); }
        }
    }

    private object CreateWorkbook(object books, string? templatePath)
    {
        activeCancellation.ThrowIfCancellationRequested();
        dynamic workbooks = books;
        if (!string.IsNullOrWhiteSpace(templatePath))
            return workbooks.Add(Path.GetFullPath(templatePath));

        try { return workbooks.Add(); }
        catch (COMException error) when (error.HResult == ExcelApplicationError)
        {
            activeCancellation.ThrowIfCancellationRequested();
            // Only the default blank-workbook path may fall back to a single worksheet.
            return workbooks.Add(XlWbatWorksheet);
        }
    }

    private void ReleaseFailedStartup(object? books)
    {
        try
        {
            object? failedApp = app;
            if (failedApp is null || books is null) return;
            dynamic instance = failedApp;
            if ((int)((dynamic)books).Count == 0 && instance.UserControl == false)
                instance.Quit();
        }
        catch (Exception) { /* Keep the workbook creation failure. */ }
        finally { Release(app); app = null; }
    }

    private void PrepareWorksheet(double rowHeight, double columnWidth, bool showGridlines)
    {
        NormalizePreservingPictures(rowHeight, columnWidth);
        activeCancellation.ThrowIfCancellationRequested();
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

    public ExcelCellPosition GetCellOrigin(int row, int column)
    {
        if (Thread.CurrentThread != thread) throw new InvalidOperationException("Excel requires its STA worker.");
        object? cells = null, cell = null, rows = null, columns = null;
        try
        {
            rows = sheet!.Rows; columns = sheet.Columns;
            if (row < 1 || row > (int)((dynamic)rows).Count || column < 1 || column > (int)((dynamic)columns).Count)
                throw new ArgumentOutOfRangeException(nameof(row), "Start cell exceeds the target worksheet limits.");
            cells = sheet.Cells; cell = ((dynamic)cells)[row, column];
            return new ExcelCellPosition((double)((dynamic)cell).Left, (double)((dynamic)cell).Top);
        }
        catch (Exception error) { throw new ExcelExportException("ExportStageWorksheet", error); }
        finally { Release(cell); Release(cells); Release(rows); Release(columns); }
    }

    private sealed record PictureGeometry(object Shape, double Left, double Top, double Width, double Height, int Placement, int LockAspectRatio);
    private static bool IsPicture(dynamic shape) => (int)shape.Type is 11 or 13; // msoLinkedPicture / msoPicture

    private void NormalizePreservingPictures(double rowHeight, double columnWidth, bool showGridlines = true)
    {
        var saved = new List<PictureGeometry>();
        object? shapes = null, rows = null, columns = null;
        var normalized = false;
        try
        {
            shapes = sheet!.Shapes;
            for (var i = 1; i <= (int)((dynamic)shapes).Count; i++)
            {
                activeCancellation.ThrowIfCancellationRequested();
                object? shape = ((dynamic)shapes).Item(i);
                try
                {
                    if (!IsPicture(shape)) continue;
                    dynamic picture = shape;
                    var geometry = new PictureGeometry(shape, (double)picture.Left, (double)picture.Top,
                        (double)picture.Width, (double)picture.Height, (int)picture.Placement, (int)picture.LockAspectRatio);
                    picture.Placement = 3; // Temporarily xlFreeFloating, only during our normalization.
                    saved.Add(geometry);
                    shape = null; // The snapshot owns this RCW until restoration.
                }
                finally { Release(shape); }
            }
            activeCancellation.ThrowIfCancellationRequested();
            rows = sheet.Rows; columns = sheet.Columns;
            ((dynamic)rows).RowHeight = rowHeight;
            activeCancellation.ThrowIfCancellationRequested();
            ((dynamic)columns).ColumnWidth = columnWidth;
            normalized = true;
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
            if (normalized && failure is not null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }

    public void ClearPictures()
    {
        if (Thread.CurrentThread != thread) throw new InvalidOperationException("Excel requires its STA worker.");
        object? shapes = null;
        try
        {
            shapes = sheet!.Shapes;
            for (var i = (int)((dynamic)shapes).Count; i >= 1; i--)
            {
                activeCancellation.ThrowIfCancellationRequested();
                object? shape = null;
                try
                {
                    shape = ((dynamic)shapes).Item(i);
                    if (IsPicture(shape)) ((dynamic)shape).Delete();
                }
                finally { Release(shape); }
            }
            nextTop = 0;
        }
        catch (Exception error) when (error is not OperationCanceledException)
        { throw new ExcelExportException("ExportStageClearPictures", error); }
        finally { Release(shapes); }
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
        catch (Exception error)
        {
            try { if (shape is not null) ((dynamic)shape).Delete(); }
            catch (Exception) { /* Keep the original Excel failure. */ }
            throw new ExcelExportException("ExportStageInsertPicture", error, Path.GetFileName(path));
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
        catch (Exception error) { throw new ExcelExportException("ExportStageSelectPictures", error); }
        finally { Release(range); Release(shapes); }
    }

    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    public void Dispose() => Dispose(waitForWorker: true);

    internal void Dispose(bool waitForWorker)
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
        if (waitForWorker) thread.Join(TimeSpan.FromSeconds(2));
    }
}
