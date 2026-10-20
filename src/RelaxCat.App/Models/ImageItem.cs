using Avalonia.Media.Imaging;
using RelaxCat.Core.Export;

namespace RelaxCat.App.Models;

public sealed class ImageItem : ObservableModel, IDisposable
{
    private bool includeInExport = true;
    private static long nextImportOrder;
    public long ImportOrder { get; } = Interlocked.Increment(ref nextImportOrder);
    public DateTime FileTimeUtc { get; }
    public ImageItem(ImportedImage image, Bitmap thumbnail, UiText text)
    { Image = image; Thumbnail = thumbnail; Text = text; FileTimeUtc = File.GetLastWriteTimeUtc(image.Path); }
    public ImportedImage Image { get; }
    public Bitmap Thumbnail { get; }
    public UiText Text { get; }
    public string Name => Path.GetFileName(Image.Path);
    public string Details => $"{Image.Metadata.PixelWidth} × {Image.Metadata.PixelHeight}  ·  {Image.Metadata.DpiX:0.#} DPI";
    public bool IncludeInExport
    {
        get => includeInExport;
        set { if (Set(ref includeInExport, value)) Notify(nameof(ThumbnailOpacity)); }
    }
    public double ThumbnailOpacity => IncludeInExport ? 1 : .38;
    public void Dispose() => Thumbnail.Dispose();
}
