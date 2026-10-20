using RelaxCat.Core.Imaging;
using RelaxCat.Core.Layout;

namespace RelaxCat.Core.Tests;

public class ScaleRuleEngineTests
{
    [Fact]
    public void First_matching_rule_wins()
    {
        var rules = """
            width > 1 and width < 1000
            scale = 0.9

            width >= 500
            scale = 0.5
            """;
        var value = new ScaleRuleEngine().Evaluate(rules, new ScaleRuleContext(800, 600, 96, 96));
        Assert.Equal(0.9, value);
    }

    [Fact]
    public void And_has_higher_precedence_than_or()
    {
        var rules = """
            width > 1000 or height > 500 and width > 700
            scale = 0.75
            """;
        Assert.Equal(0.75, new ScaleRuleEngine().Evaluate(rules, new ScaleRuleContext(800, 600, 96, 96)));
    }
}

public class ImageSizingTests
{
    [Fact]
    public void Dpi_is_normalized_before_display_sizing()
    {
        var options = new SizingOptions(1920, 1080, 0.5, 2);
        var a = ImageSizing.Calculate(new ImageMetadata(960, 540, 96, 96), options);
        var b = ImageSizing.Calculate(new ImageMetadata(1920, 1080, 192, 192), options);
        Assert.Equal(a.DisplayWidth, b.DisplayWidth, 6);
        Assert.Equal(a.DisplayHeight, b.DisplayHeight, 6);
    }

    [Fact]
    public void Fallback_respects_maximum_enlargement()
    {
        var result = ImageSizing.Calculate(new ImageMetadata(100, 100), new SizingOptions(1920, 1080, 1, 1.5));
        Assert.Equal(1.5, result.Scale);
    }
}

public class ImageLayoutEngineTests
{
    [Fact]
    public void Space_saving_wraps_when_next_image_does_not_fit()
    {
        var images = new[] { new LayoutImage("a", 400, 200), new LayoutImage("b", 300, 300), new LayoutImage("c", 350, 100) };
        var result = new ImageLayoutEngine().Layout(images, new LayoutOptions(10, 20, 20, 30, 1000, true));
        Assert.Equal(0, result[0].Row);
        Assert.Equal(0, result[1].Row);
        Assert.Equal(1, result[2].Row);
        Assert.Equal(350, result[2].Y); // 20 start + 300 tallest + 30 gap
    }

    [Fact]
    public void Default_mode_places_one_image_per_row()
    {
        var images = new[] { new LayoutImage("a", 100, 100), new LayoutImage("b", 100, 100) };
        var result = new ImageLayoutEngine().Layout(images, new LayoutOptions(0, 0, 10, 20, 1000, false));
        Assert.Equal(0, result[0].Row);
        Assert.Equal(1, result[1].Row);
        Assert.Equal(120, result[1].Y);
    }
}
