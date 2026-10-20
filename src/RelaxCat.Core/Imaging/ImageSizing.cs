namespace RelaxCat.Core.Imaging;

public readonly record struct ImageMetadata(int PixelWidth, int PixelHeight, double DpiX = 96, double DpiY = 96)
{
    public double AspectRatio => PixelHeight == 0 ? 0 : (double)PixelWidth / PixelHeight;
}

public sealed record SizingOptions(
    double DesignWidth,
    double DesignHeight,
    double DesignScale,
    double MaximumEnlargement,
    string Rules = "");

public readonly record struct ScaleResult(
    double Scale,
    double DisplayWidth,
    double DisplayHeight,
    bool RuleMatched);

public static class ImageSizing
{
    public const double ReferenceDpi = 96d;

    public static ScaleResult Calculate(ImageMetadata image, SizingOptions options, ScaleRuleEngine? ruleEngine = null)
    {
        if (image.PixelWidth <= 0 || image.PixelHeight <= 0)
            throw new ArgumentOutOfRangeException(nameof(image), "Image dimensions must be positive.");
        if (options.DesignWidth <= 0 || options.DesignHeight <= 0 || options.DesignScale <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Design dimensions and scale must be positive.");
        if (options.MaximumEnlargement <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum enlargement must be positive.");

        var dpiX = image.DpiX > 0 ? image.DpiX : ReferenceDpi;
        var dpiY = image.DpiY > 0 ? image.DpiY : ReferenceDpi;
        var normalizedWidth = image.PixelWidth * ReferenceDpi / dpiX;
        var normalizedHeight = image.PixelHeight * ReferenceDpi / dpiY;

        double scale;
        var matched = false;
        if (ruleEngine is not null && !string.IsNullOrWhiteSpace(options.Rules))
        {
            var context = new ScaleRuleContext(image.PixelWidth, image.PixelHeight, dpiX, dpiY);
            var ruleScale = ruleEngine.Evaluate(options.Rules, context);
            if (ruleScale is > 0)
            {
                scale = Math.Min(ruleScale.Value, options.MaximumEnlargement);
                matched = true;
            }
            else
            {
                scale = CalculateFallback(normalizedWidth, normalizedHeight, options);
            }
        }
        else
        {
            scale = CalculateFallback(normalizedWidth, normalizedHeight, options);
        }

        return new ScaleResult(scale, normalizedWidth * scale, normalizedHeight * scale, matched);
    }

    private static double CalculateFallback(double width, double height, SizingOptions options)
    {
        var targetWidth = options.DesignWidth * options.DesignScale;
        var targetHeight = options.DesignHeight * options.DesignScale;
        return Math.Min(Math.Min(targetWidth / width, targetHeight / height), options.MaximumEnlargement);
    }
}
