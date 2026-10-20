using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;
using RelaxCat.Core.Layout;

namespace RelaxCat.Core.Tests;

public class WorkflowTests
{
    private static readonly ImportedImage[] Images = {
        new("a.png", new ImageMetadata(400, 200)), new("b.png", new ImageMetadata(400, 200)) };
    private static readonly ExportOptions Options = new(new SizingOptions(1000, 1000, 1, 1));

    [Fact]
    public void Export_normalizes_before_inserting_and_converts_pixels_to_points()
    {
        var target = new FakeTarget { NextTop = 200 };
        new ImageExporter().Export(Images, Options, target, null, default);
        Assert.Equal(new[] { "prepare", "insert", "insert" }, target.Events);
        Assert.Equal(300, target.Placements[0].Width);
        Assert.Equal(150, target.Placements[0].Height);
        Assert.Equal(212, target.Placements[0].Y);
        Assert.Equal(374, target.Placements[1].Y);
    }

    [Fact]
    public void Invalid_later_rule_is_rejected_before_workbook_is_modified()
    {
        var target = new FakeTarget();
        var sizing = Options.Sizing with { Rules = "width > 0\nscale = 1\nwidth > 0 or invalid\nscale = 2" };
        Assert.Throws<ScaleRuleException>(() => new ImageExporter().Export(Images, Options with { Sizing = sizing }, target, null, default));
        Assert.Empty(target.Events);
    }

    [Fact]
    public void Cancellation_keeps_completed_images_and_stops_before_the_next()
    {
        using var cancellation = new CancellationTokenSource();
        var target = new FakeTarget { AfterInsert = cancellation.Cancel };
        Assert.Throws<OperationCanceledException>(() => new ImageExporter().Export(Images, Options, target, null, cancellation.Token));
        Assert.Single(target.Placements);
    }

    [Fact]
    public void Precancelled_export_does_not_touch_excel()
    {
        var target = new FakeTarget();
        Assert.Throws<OperationCanceledException>(() => new ImageExporter().Export(Images, Options, target, null, new CancellationToken(true)));
        Assert.Empty(target.Events);
    }

    [Fact]
    public void Rule_scale_is_used_without_fallback_enlargement_limit()
    {
        var result = ImageSizing.Calculate(new ImageMetadata(100, 100), new SizingOptions(1000, 1000, 1, 1.5, "width > 0\nscale = 3"));
        Assert.Equal(3, result.Scale);
        Assert.True(result.RuleMatched);
    }

    [Fact]
    public void Cache_distinguishes_rules_dpi_and_design_configuration()
    {
        var cache = new ScaleCache(); var engine = new ScaleRuleEngine();
        var metadata = Images[0].Metadata;
        var options = Options.Sizing;
        var first = cache.GetOrCalculate(metadata, options, engine);
        Assert.Equal(first, cache.GetOrCalculate(metadata, options, engine));
        cache.GetOrCalculate(metadata with { DpiX = 192 }, options, engine);
        cache.GetOrCalculate(metadata, options with { DesignScale = .5 }, engine);
        cache.GetOrCalculate(metadata, options with { Rules = "width > 0\nscale = 0.5" }, engine);
        Assert.Equal(4, cache.Count);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(-1)]
    public void Invalid_sizes_are_rejected(double value)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ImageSizing.Calculate(Images[0].Metadata, Options.Sizing with { DesignWidth = value }));
    }

    [Fact]
    public void Compact_export_wraps_using_design_width_in_points()
    {
        var target = new FakeTarget();
        new ImageExporter().Export(Images, Options with { SpaceSaving = true }, target, null, default);
        Assert.Equal(target.Placements[0].Y, target.Placements[1].Y);
        Assert.Equal(324, target.Placements[1].X);
    }

    private sealed class FakeTarget : IExcelTarget
    {
        public List<string> Events { get; } = new();
        public List<ImagePlacement> Placements { get; } = new();
        public Action? AfterInsert { get; init; }
        public double NextTop { get; init; }
        public ExcelCellSize CellSize { get; init; } = new(60, 18);
        public void SelectAllPictures() => Events.Add("select");
        public void Prepare(string? path, double row, double column) => Events.Add("prepare");
        public void Insert(string path, ImagePlacement placement)
        { Events.Add("insert"); Placements.Add(placement); AfterInsert?.Invoke(); }
    }
}
