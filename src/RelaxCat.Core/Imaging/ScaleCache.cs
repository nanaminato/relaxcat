using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace RelaxCat.Core.Imaging;

public sealed class ScaleCache
{
    private readonly ConcurrentDictionary<ScaleCacheKey, ScaleResult> _cache = new();

    public ScaleResult GetOrCalculate(ImageMetadata image, SizingOptions options, ScaleRuleEngine engine)
    {
        var key = ScaleCacheKey.Create(image, options);
        return _cache.GetOrAdd(key, _ => ImageSizing.Calculate(image, options, engine));
    }

    public void Clear() => _cache.Clear();
    public int Count => _cache.Count;
}

public readonly record struct ScaleCacheKey(
    int Width, int Height, double DpiX, double DpiY,
    double DesignWidth, double DesignHeight, double DesignScale, double MaximumEnlargement,
    string RuleHash)
{
    public static ScaleCacheKey Create(ImageMetadata image, SizingOptions options) => new(
        image.PixelWidth, image.PixelHeight, image.DpiX, image.DpiY,
        options.DesignWidth, options.DesignHeight, options.DesignScale, options.MaximumEnlargement,
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(options.Rules ?? string.Empty))));
}
