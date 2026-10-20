using System.Globalization;
using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;

namespace RelaxCat.App.Models;

public sealed class SettingsModel : ObservableModel
{
    private string designWidth = "1920", designHeight = "1080", designScale = "0.5", maximumEnlargement = "1.5";
    private string horizontalGap = "12", verticalGap = "12", rowHeight = "18", columnWidth = "10", templatePath = "", rules = "";
    private bool compact;
    private bool showGridlines = true;
    private bool selectAllPicturesAfterExport;
    private int horizontalGapUnitIndex, verticalGapUnitIndex;
    public string DesignWidth { get => designWidth; set => Set(ref designWidth, value); }
    public string DesignHeight { get => designHeight; set => Set(ref designHeight, value); }
    public string DesignScale { get => designScale; set => Set(ref designScale, value); }
    public string MaximumEnlargement { get => maximumEnlargement; set => Set(ref maximumEnlargement, value); }
    public string HorizontalGap { get => horizontalGap; set => Set(ref horizontalGap, value); }
    public string VerticalGap { get => verticalGap; set => Set(ref verticalGap, value); }
    public string RowHeight { get => rowHeight; set => Set(ref rowHeight, value); }
    public string ColumnWidth { get => columnWidth; set => Set(ref columnWidth, value); }
    public string TemplatePath { get => templatePath; set => Set(ref templatePath, value); }
    public string Rules { get => rules; set => Set(ref rules, value); }
    public bool ShowGridlines { get => showGridlines; set => Set(ref showGridlines, value); }
    public bool Compact { get => compact; set => Set(ref compact, value); }
    public bool SelectAllPicturesAfterExport { get => selectAllPicturesAfterExport; set => Set(ref selectAllPicturesAfterExport, value); }
    public int HorizontalGapUnitIndex { get => horizontalGapUnitIndex; set { if (value is 0 or 1) Set(ref horizontalGapUnitIndex, value); } }
    public int VerticalGapUnitIndex { get => verticalGapUnitIndex; set { if (value is 0 or 1) Set(ref verticalGapUnitIndex, value); } }

    public ExportOptions CreateOptions(UiText text)
    {
        double Number(string value, bool zero = false)
        {
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ||
                !double.IsFinite(n) || (zero ? n < 0 : n <= 0)) throw new ArgumentException(text["NumberError"]);
            return n;
        }
        var sizing = new SizingOptions(Number(DesignWidth), Number(DesignHeight), Number(DesignScale), Number(MaximumEnlargement), Rules);
        var result = new ExportOptions(sizing, Compact, Number(HorizontalGap, true), Number(VerticalGap, true), Number(RowHeight), Number(ColumnWidth), TemplatePath,
            (GapUnit)HorizontalGapUnitIndex, (GapUnit)VerticalGapUnitIndex, SelectAllPicturesAfterExport, ShowGridlines);
        if (result.RowHeight > 409 || result.ColumnWidth > 255) throw new ArgumentException(text["ExcelRangeError"]);
        new ScaleRuleEngine().ParseRules(Rules);
        ImageSizing.Calculate(new ImageMetadata(1, 1), sizing);
        if (!string.IsNullOrWhiteSpace(TemplatePath) && !File.Exists(TemplatePath)) throw new FileNotFoundException(TemplatePath);
        return result;
    }
}
