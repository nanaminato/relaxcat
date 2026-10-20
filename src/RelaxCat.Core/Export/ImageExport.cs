using RelaxCat.Core.Imaging;
using RelaxCat.Core.Layout;

namespace RelaxCat.Core.Export;

public sealed record ImportedImage(string Path, ImageMetadata Metadata);
public enum GapUnit { Points, Cells }
public readonly record struct ExcelCellSize(double WidthPoints, double HeightPoints);
public sealed record ExportOptions(SizingOptions Sizing, bool SpaceSaving = false,
    double HorizontalGap = 12, double VerticalGap = 12, double RowHeight = 18,
    double ColumnWidth = 10, string? TemplatePath = null,
    GapUnit HorizontalGapUnit = GapUnit.Points, GapUnit VerticalGapUnit = GapUnit.Points,
    bool SelectAllPicturesAfterExport = false);

public interface IExcelTarget
{
    void Prepare(string? templatePath, double rowHeight, double columnWidth);
    // Bottom edge of existing worksheet content in points; zero for an empty sheet.
    double NextTop { get; }
    // Actual physical cell dimensions after Prepare, not Excel character units.
    ExcelCellSize CellSize { get; }
    void Insert(string path, ImagePlacement points);
    void SelectAllPictures();
}

public sealed class ImageExporter
{
    private readonly ScaleCache cache = new();
    private readonly ScaleRuleEngine rules = new();

    public int Export(IReadOnlyList<ImportedImage> images, ExportOptions options, IExcelTarget target,
        IProgress<int>? progress, CancellationToken cancellationToken)
    {
        if (images.Count == 0) return 0;
        // Validate source sizes/configuration first; cell dimensions require a normalized worksheet.
        var sizes = images.Select((image, index) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            var size = cache.GetOrCalculate(image.Metadata, options.Sizing, rules);
            return new LayoutImage(index.ToString(), size.DisplayWidth * .75, size.DisplayHeight * .75);
        }).ToArray();
        var layoutOptions = new LayoutOptions(12, 12, options.HorizontalGap, options.VerticalGap,
            options.Sizing.DesignWidth * options.Sizing.DesignScale * .75, options.SpaceSaving);
        var layout = new ImageLayoutEngine();
        layout.Layout(sizes, layoutOptions); // Validate before touching a live worksheet.
        if (!Enum.IsDefined(options.HorizontalGapUnit) || !Enum.IsDefined(options.VerticalGapUnit))
            throw new ArgumentOutOfRangeException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        target.Prepare(options.TemplatePath, options.RowHeight, options.ColumnWidth);
        var horizontalGap = ResolveGap(options.HorizontalGap, options.HorizontalGapUnit, target.CellSize.WidthPoints);
        var verticalGap = ResolveGap(options.VerticalGap, options.VerticalGapUnit, target.CellSize.HeightPoints);
        var placements = layout.Layout(sizes, layoutOptions with
        {
            HorizontalGap = horizontalGap, VerticalGap = verticalGap,
            StartY = target.NextTop > 0 ? target.NextTop + verticalGap : 12
        });
        var completed = 0;
        foreach (var placement in placements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            target.Insert(images[completed].Path, placement);
            progress?.Report(++completed);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (options.SelectAllPicturesAfterExport) target.SelectAllPictures();
        return completed;
    }

    private static double ResolveGap(double value, GapUnit unit, double cellPoints)
    {
        if (unit == GapUnit.Points) return value;
        if (!double.IsFinite(cellPoints) || cellPoints <= 0) throw new InvalidOperationException("Invalid worksheet cell dimensions.");
        var result = value * cellPoints;
        if (!double.IsFinite(result)) throw new ArgumentOutOfRangeException(nameof(value));
        return result;
    }
}
