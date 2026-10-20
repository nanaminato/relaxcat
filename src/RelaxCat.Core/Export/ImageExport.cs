using RelaxCat.Core.Imaging;
using RelaxCat.Core.Layout;

namespace RelaxCat.Core.Export;

public sealed record ImportedImage(string Path, ImageMetadata Metadata);
public sealed record ExportOptions(SizingOptions Sizing, bool SpaceSaving = false,
    double HorizontalGap = 12, double VerticalGap = 12, double RowHeight = 18,
    double ColumnWidth = 10, string? TemplatePath = null);

public interface IExcelTarget
{
    void Prepare(string? templatePath, double rowHeight, double columnWidth);
    double NextTop { get; }
    void Insert(string path, ImagePlacement points);
}

public sealed class ImageExporter
{
    private readonly ScaleCache cache = new();
    private readonly ScaleRuleEngine rules = new();

    public int Export(IReadOnlyList<ImportedImage> images, ExportOptions options, IExcelTarget target,
        IProgress<int>? progress, CancellationToken cancellationToken)
    {
        if (images.Count == 0) return 0;
        // Validate and plan everything before modifying a live worksheet.
        var sizes = images.Select((image, index) =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            var size = cache.GetOrCalculate(image.Metadata, options.Sizing, rules);
            return new LayoutImage(index.ToString(), size.DisplayWidth * .75, size.DisplayHeight * .75);
        }).ToArray();
        var layoutOptions = new LayoutOptions(12, 12, options.HorizontalGap, options.VerticalGap,
            options.Sizing.DesignWidth * options.Sizing.DesignScale * .75, options.SpaceSaving);
        var placements = new ImageLayoutEngine().Layout(sizes, layoutOptions);
        cancellationToken.ThrowIfCancellationRequested();
        target.Prepare(options.TemplatePath, options.RowHeight, options.ColumnWidth);
        var top = target.NextTop;
        var completed = 0;
        foreach (var placement in placements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            target.Insert(images[completed].Path, placement with { Y = placement.Y + top });
            progress?.Report(++completed);
        }
        return completed;
    }
}
