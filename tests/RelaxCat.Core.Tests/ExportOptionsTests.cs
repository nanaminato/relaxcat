using RelaxCat.Core.Export;
using RelaxCat.Core.Imaging;
using RelaxCat.Core.Layout;

namespace RelaxCat.Core.Tests;

public class ExportOptionsTests
{
    private static readonly ImportedImage[] Images = {
        new("a.png", new ImageMetadata(400, 200)), new("b.png", new ImageMetadata(400, 200)) };
    private static readonly ExportOptions Options = new(new SizingOptions(1000, 1000, 1, 1));

    [Theory]
    [InlineData(GapUnit.Points, 2)]
    [InlineData(GapUnit.Cells, 135)]
    public void Horizontal_gap_uses_requested_unit_and_actual_column_width(GapUnit unit, double gap)
    {
        var target = new Target();
        new ImageExporter().Export(Images, Options with { SpaceSaving = true, HorizontalGap = 2, HorizontalGapUnit = unit }, target, null, default);
        Assert.Equal(gap, target.Placements[1].X - target.Placements[0].X - target.Placements[0].Width, 8);
    }

    [Theory]
    [InlineData(GapUnit.Points, 1.5)]
    [InlineData(GapUnit.Cells, 36)]
    public void Vertical_gap_uses_requested_unit_and_actual_row_height(GapUnit unit, double gap)
    {
        var target = new Target();
        new ImageExporter().Export(Images, Options with { VerticalGap = 1.5, VerticalGapUnit = unit }, target, null, default);
        Assert.Equal(gap, target.Placements[1].Y - target.Placements[0].Y - target.Placements[0].Height, 8);
    }

    [Fact]
    public void Fractional_cell_gaps_and_append_spacing_are_supported()
    {
        var target = new Target { NextTop = 500 };
        new ImageExporter().Export(Images, Options with { SpaceSaving = true, HorizontalGap = .5, VerticalGap = 2,
            HorizontalGapUnit = GapUnit.Cells, VerticalGapUnit = GapUnit.Cells }, target, null, default);
        Assert.Equal(548, target.Placements[0].Y);
        Assert.Equal(345.75, target.Placements[1].X);
    }

    [Fact]
    public void Selection_runs_only_after_all_images_are_inserted_when_enabled()
    {
        var target = new Target();
        new ImageExporter().Export(Images, Options with { SelectAllPicturesAfterExport = true }, target, null, default);
        Assert.Equal(new[] { "prepare", "insert", "insert", "select" }, target.Events);
    }

    [Fact]
    public void Selection_is_opt_in()
    {
        var target = new Target();
        new ImageExporter().Export(Images, Options, target, null, default);
        Assert.DoesNotContain("select", target.Events);
    }

    [Fact]
    public void Cancelled_export_does_not_select_pictures()
    {
        using var cancellation = new CancellationTokenSource();
        var target = new Target { AfterInsert = cancellation.Cancel };
        Assert.Throws<OperationCanceledException>(() => new ImageExporter().Export(Images,
            Options with { SelectAllPicturesAfterExport = true }, target, null, cancellation.Token));
        Assert.DoesNotContain("select", target.Events);
    }

    [Fact]
    public void Failed_export_does_not_select_pictures()
    {
        var target = new Target { AfterInsert = () => throw new IOException() };
        Assert.Throws<IOException>(() => new ImageExporter().Export(Images,
            Options with { SelectAllPicturesAfterExport = true }, target, null, default));
        Assert.DoesNotContain("select", target.Events);
    }

    [Fact]
    public void Empty_export_does_not_create_or_select_a_workbook()
    {
        var target = new Target();
        Assert.Equal(0, new ImageExporter().Export(Array.Empty<ImportedImage>(), Options with { SelectAllPicturesAfterExport = true }, target, null, default));
        Assert.Empty(target.Events);
    }

    [Fact]
    public void Invalid_gap_units_are_rejected_before_excel_is_modified()
    {
        var target = new Target();
        Assert.Throws<ArgumentOutOfRangeException>(() => new ImageExporter().Export(Images,
            Options with { HorizontalGapUnit = (GapUnit)17 }, target, null, default));
        Assert.Empty(target.Events);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Gridline_preference_reaches_target_on_every_export(bool visible)
    {
        var target = new Target();
        var exporter = new ImageExporter();
        exporter.Export(Images, Options with { ShowGridlines = visible }, target, null, default);
        Assert.Equal(visible, target.Gridlines);
        exporter.Export(Images, Options with { ShowGridlines = !visible }, target, null, default);
        Assert.Equal(!visible, target.Gridlines);
    }

    private sealed class Target : IExcelTarget
    {
        public List<string> Events { get; } = new();
        public List<ImagePlacement> Placements { get; } = new();
        public double NextTop { get; init; }
        public Action? AfterInsert { get; init; }
        public ExcelCellSize CellSize
        {
            get { Assert.Equal("prepare", Events[0]); return new ExcelCellSize(67.5, 24); }
        }
        public bool Gridlines { get; private set; }
        public void Prepare(string? templatePath, double rowHeight, double columnWidth, bool showGridlines = true)
        { Gridlines = showGridlines; Events.Add("prepare"); }
        public void Insert(string path, ImagePlacement points) { Events.Add("insert"); Placements.Add(points); AfterInsert?.Invoke(); }
        public void SelectAllPictures() => Events.Add("select");
    }
}
