namespace RelaxCat.Core.Layout;

public readonly record struct LayoutImage(string Id, double Width, double Height);
public readonly record struct ImagePlacement(string Id, double X, double Y, double Width, double Height, int Row);

public sealed record LayoutOptions(
    double StartX,
    double StartY,
    double HorizontalGap,
    double VerticalGap,
    double MaximumRowWidth,
    bool SpaceSaving);

public sealed class ImageLayoutEngine
{
    public IReadOnlyList<ImagePlacement> Layout(IEnumerable<LayoutImage> images, LayoutOptions options)
    {
        if (options.MaximumRowWidth <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaximumRowWidth));

        var result = new List<ImagePlacement>();
        var x = options.StartX;
        var y = options.StartY;
        var rowHeight = 0d;
        var row = 0;
        var rowHasImage = false;

        foreach (var image in images)
        {
            if (image.Width <= 0 || image.Height <= 0) throw new ArgumentOutOfRangeException(nameof(images), "Image dimensions must be positive.");

            var mustWrap = rowHasImage && (!options.SpaceSaving || x - options.StartX + options.HorizontalGap + image.Width > options.MaximumRowWidth);
            if (mustWrap)
            {
                y += rowHeight + options.VerticalGap;
                x = options.StartX;
                rowHeight = 0;
                row++;
                rowHasImage = false;
            }

            if (rowHasImage) x += options.HorizontalGap;
            result.Add(new ImagePlacement(image.Id, x, y, image.Width, image.Height, row));
            x += image.Width;
            rowHeight = Math.Max(rowHeight, image.Height);
            rowHasImage = true;
        }

        return result;
    }
}
