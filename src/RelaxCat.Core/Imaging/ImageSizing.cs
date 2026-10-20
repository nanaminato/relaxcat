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
        if (!double.IsFinite(options.DesignWidth) || !double.IsFinite(options.DesignHeight) || !double.IsFinite(options.DesignScale) || options.DesignWidth <= 0 || options.DesignHeight <= 0 || options.DesignScale <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Design dimensions and scale must be positive.");
        if (!double.IsFinite(options.MaximumEnlargement) || options.MaximumEnlargement <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum enlargement must be positive.");

        var dpiX = double.IsFinite(image.DpiX) && image.DpiX > 0 ? image.DpiX : ReferenceDpi;
        var dpiY = double.IsFinite(image.DpiY) && image.DpiY > 0 ? image.DpiY : ReferenceDpi;
        // DPI controls uniform physical size, never the source pixel aspect ratio.
        var effectiveDpi = dpiX / 2d + dpiY / 2d;
        var normalizedWidth = image.PixelWidth * ReferenceDpi / effectiveDpi;
        var normalizedHeight = image.PixelHeight * ReferenceDpi / effectiveDpi;

        double scale;
        var matched = false;
        if (!string.IsNullOrWhiteSpace(options.Rules))
        {
            ruleEngine ??= new ScaleRuleEngine();
            var context = new ScaleRuleContext(image.PixelWidth, image.PixelHeight, dpiX, dpiY);
            var ruleScale = ruleEngine.Evaluate(options.Rules, context);
            if (ruleScale is > 0)
            {
                scale = ruleScale.Value;
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

        var displayWidth = normalizedWidth * scale;
        var displayHeight = normalizedHeight * scale;
        if (!double.IsFinite(scale) || !double.IsFinite(displayWidth) || !double.IsFinite(displayHeight) || displayWidth <= 0 || displayHeight <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Calculated display dimensions must be positive and finite.");
        return new ScaleResult(scale, displayWidth, displayHeight, matched);
    }

    private static double CalculateFallback(double width, double height, SizingOptions options)
    {
        var targetWidth = options.DesignWidth * options.DesignScale;
        var targetHeight = options.DesignHeight * options.DesignScale;
        return Math.Min(Math.Min(targetWidth / width, targetHeight / height), options.MaximumEnlargement);
    }
}
