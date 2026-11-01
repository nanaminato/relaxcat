using System.Globalization;
using System.Text;
using RelaxCat.Core.Excel;
using Xunit;

namespace RelaxCat.Core.Tests;

public sealed class ExcelStyleTests : IDisposable
{
    private const string DefaultXf = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""";
    private const string RedXf = """<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>""";
    private const string BorderXf = """<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>""";
    private const string DateXf = """<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>""";

    private const string Fonts = """
        <fonts count="2">
          <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
          <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>
        </fonts>
        """;

    private const string Fills = """
        <fills count="3">
          <fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="gray125"/></fill>
          <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>
        </fills>
        """;

    private const string Borders = """
        <borders count="2">
          <border><left/><right/><top/><bottom/><diagonal/></border>
          <border><left/><right/><top/><bottom style="thin"><color rgb="FFFFCC00"/></bottom><diagonal/></border>
        </borders>
        """;

    private const string NumberFormats = """
        <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
        """;

    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-excel-style-" + Guid.NewGuid().ToString("N"));

    public ExcelStyleTests() => Directory.CreateDirectory(directory);

    public void Dispose()
    {
        try { Directory.Delete(directory, true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private string PathOf(string name) => Path.Combine(directory, name);

    private static string SampleStyles() => OoxmlWorkbookFixture.StylesXml(
        OoxmlWorkbookFixture.CellXfs(DefaultXf, RedXf, RedXf, BorderXf, DateXf),
        Fonts, Fills, Borders, NumberFormats);

    private string CreateSample(string name = "sample.xlsx") => OoxmlWorkbookFixture.Create(PathOf(name), SampleStyles(),
        new SheetSpec("Data", DataSheet()),
        new SheetSpec("Hidden", OoxmlWorkbookFixture.Worksheet("A1",
            """<row r="1"><c r="A1"><v>1</v></c></row>"""), true));

    private static string DataSheet() => OoxmlWorkbookFixture.Worksheet("A1:C3", """
        <row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>
        <row r="2"><c r="A2" s="2"><v>3</v></c></row>
        <row r="3" s="3"><c r="C3"><v>4</v></c></row>
        """);

    [Fact]
    public void Analyze_CountsStylesAndPreviewsWorksheetFormats()
    {
        var input = CreateSample();
        var stylesXml = SampleStyles();

        var analysis = new ExcelStyleAnalyzer().Analyze(input);

        Assert.Equal(2, analysis.FontCount);
        Assert.Equal(3, analysis.FillCount);
        Assert.Equal(2, analysis.BorderCount);
        Assert.Equal(1, analysis.CellStyleXfCount);
        Assert.Equal(5, analysis.CellXfCount);
        Assert.Equal(4, analysis.UniqueCellXfCount);
        Assert.Equal(1, analysis.DuplicateCellXfCount);
        Assert.Equal(4, analysis.ReferencedCellXfCount);
        Assert.Equal(1, analysis.UnreferencedCellXfCount);
        Assert.Equal(2, analysis.SheetCount);
        Assert.False(analysis.MacroEnabled);
        Assert.True(analysis.HasDuplicates);
        Assert.Equal(0.2, analysis.ReductionRatio, 3);
        Assert.Equal(Encoding.UTF8.GetByteCount(stylesXml), analysis.StylesXmlBytes);
        Assert.True(analysis.ProjectedStylesXmlBytes < analysis.StylesXmlBytes);

        var data = analysis.Worksheets.Single(sheet => sheet.Name == "Data");
        Assert.Equal("A1:C3", data.UsedRange);
        Assert.Equal(4, data.CellCount);
        Assert.Equal(4, data.DistinctStyleCount);
        Assert.False(data.Hidden);
        Assert.True(analysis.Worksheets.Single(sheet => sheet.Name == "Hidden").Hidden);

        var red = data.Styles.Single(usage => usage.StyleId == 1).Style;
        Assert.Equal("#FF0000", red.FillColor);
        Assert.Equal("Arial", red.FontName);
        Assert.Equal(14, red.FontSize);
        Assert.True(red.Bold);
        Assert.Equal("#FF0000", red.FontColor);
        Assert.False(red.IsDefault);
        Assert.Equal(1, data.Styles.Single(usage => usage.StyleId == 1).CellCount);
        Assert.Equal(red with { StyleId = 2 }, data.Styles.Single(usage => usage.StyleId == 2).Style); // identical XML resolves identically

        var border = data.Styles.Single(usage => usage.StyleId == 3).Style;
        Assert.Equal(BorderSides.Bottom, border.Borders);
        Assert.Equal("#FFCC00", border.BorderColor);
        Assert.Equal(.9, border.BorderThickness);
        Assert.Null(border.FillColor);
        Assert.Equal(1, data.Styles.Single(usage => usage.StyleId == 3).CellCount);

        var normal = data.Styles.Single(usage => usage.StyleId == 0).Style;
        Assert.True(normal.IsDefault);
        Assert.Equal(1, data.Styles.Single(usage => usage.StyleId == 0).CellCount);
    }

    [Fact]
    public void Optimize_MergesIdenticalFormatsWithoutChangingVisualStyles()
    {
        var input = CreateSample();
        var output = PathOf("optimized.xlsx");
        var originalBytes = File.ReadAllBytes(input);

        var result = new ExcelStyleOptimizer().Optimize(input, output);

        Assert.True(result.Verified);
        Assert.All(result.Checks, check => Assert.True(check.Passed, check.Key + ": " + check.Detail));
        Assert.Equal(5, result.InputCellXfCount);
        Assert.Equal(4, result.OutputCellXfCount);
        Assert.Equal(1, result.RemovedCellXfCount);
        Assert.Equal(2, result.RemappedReferences); // A2's duplicate format and the row-level border format
        Assert.Equal(2, result.WorksheetCount);
        Assert.Equal(originalBytes, File.ReadAllBytes(input));
        Assert.True(File.Exists(output));

        var optimized = new ExcelStyleAnalyzer().Analyze(output);
        Assert.Equal(0, optimized.DuplicateCellXfCount);
        Assert.Equal(4, optimized.CellXfCount);
        Assert.Equal(3, optimized.ReferencedCellXfCount);
        Assert.Equal(1, optimized.UnreferencedCellXfCount); // unique, unreferenced formats are kept
        Assert.Equal(optimized.StylesXmlBytes, result.OutputStylesXmlBytes);
        Assert.Equal(new FileInfo(output).Length, result.OutputFileSize);
        Assert.Equal(result.InputFileSize, new FileInfo(input).Length);

        AssertSameVisualStyles(new ExcelStyleAnalyzer().Analyze(input), optimized);
        Assert.Equal(OoxmlWorkbookFixture.PartNames(input), OoxmlWorkbookFixture.PartNames(output));
    }

    [Fact]
    public void Optimize_PreservesMacroAndBinaryParts()
    {
        var input = CreateSample("macro.xlsm");
        var payload = new byte[] { 1, 2, 3, 4, 5, 6, 7 };
        OoxmlWorkbookFixture.AddPart(input, "xl/vbaProject.bin", payload);
        var output = PathOf("macro-optimized.xlsm");
        Assert.True(new ExcelStyleAnalyzer().Analyze(input).MacroEnabled);

        var result = new ExcelStyleOptimizer().Optimize(input, output);

        Assert.True(result.Verified);
        Assert.True(new ExcelStyleAnalyzer().Analyze(output).MacroEnabled);
        Assert.Equal(payload, OoxmlWorkbookFixture.ReadPart(output, "xl/vbaProject.bin"));
        Assert.Throws<ArgumentException>(() => new ExcelStyleOptimizer().Optimize(input, PathOf("macro.xlsx")));
    }

    [Fact]
    public void Optimize_RemapsColumnStylesAndKeepsInheritedCellFormatting()
    {
        var cellXfs = OoxmlWorkbookFixture.CellXfs(DefaultXf, RedXf, RedXf);
        var styles = OoxmlWorkbookFixture.StylesXml(cellXfs, Fonts, Fills, Borders, NumberFormats);
        var sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:B1"/>
              <cols><col min="2" max="2" style="2" width="18"/></cols>
              <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>
            </worksheet>
            """;
        var input = OoxmlWorkbookFixture.Create(PathOf("column-styles.xlsx"), styles, new SheetSpec("Sheet1", sheet));
        var output = PathOf("column-styles-optimized.xlsx");

        var result = new ExcelStyleOptimizer().Optimize(input, output);

        Assert.True(result.Verified);
        Assert.Equal(1, result.RemappedReferences);
        var optimized = new ExcelStyleAnalyzer().Analyze(output);
        var summary = optimized.Worksheets.Single();
        Assert.Equal(2, summary.CellCount);
        Assert.Equal(2, summary.DistinctStyleCount);
        Assert.Equal(1, summary.Styles.Single(usage => usage.StyleId == 1).CellCount);
        Assert.Equal("#FF0000", summary.Styles.Single(usage => usage.StyleId == 1).Style.FillColor);
    }

    [Fact]
    public void Optimize_WithoutDuplicatesStillWritesAVerifiedCopy()
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(DefaultXf,
            """<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"""));
        var input = OoxmlWorkbookFixture.Create(PathOf("clean.xlsx"), styles,
            new SheetSpec("Sheet1", OoxmlWorkbookFixture.Worksheet("A1:B1",
                """<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>""")));
        var output = PathOf("clean-copy.xlsx");
        var clean = new ExcelStyleAnalyzer().Analyze(input);
        Assert.Equal(clean.StylesXmlBytes, clean.ProjectedStylesXmlBytes);

        var result = new ExcelStyleOptimizer().Optimize(input, output);

        Assert.True(result.Verified);
        Assert.Equal(0, result.RemovedCellXfCount);
        Assert.Equal(0, result.RemappedReferences);
        Assert.True(File.Exists(output));
        Assert.Equal(0, new ExcelStyleAnalyzer().Analyze(output).DuplicateCellXfCount);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, "xl/styles.xml"),
            OoxmlWorkbookFixture.ReadPart(output, "xl/styles.xml"));
    }

    [Fact]
    public void Optimize_RejectsUnsafeTargets()
    {
        var input = CreateSample();
        var optimizer = new ExcelStyleOptimizer();

        Assert.Throws<ArgumentException>(() => optimizer.Optimize(input, input));
        Assert.Throws<ArgumentException>(() => optimizer.Optimize(input, PathOf("copy.xlsm")));
        Assert.Throws<FileNotFoundException>(() => optimizer.Optimize(PathOf("missing.xlsx"), PathOf("out.xlsx")));
        Assert.Throws<DirectoryNotFoundException>(() =>
            optimizer.Optimize(input, Path.Combine(directory, "no-such-folder", "out.xlsx")));

        var legacy = PathOf("legacy.xls");
        File.WriteAllText(legacy, "not a package");
        Assert.Throws<NotSupportedException>(() => optimizer.Optimize(legacy, PathOf("legacy-out.xls")));

        var broken = PathOf("broken.xlsx");
        File.WriteAllText(broken, "not a zip package");
        Assert.Throws<InvalidDataException>(() => new ExcelStyleAnalyzer().Analyze(broken));
        Assert.Throws<InvalidDataException>(() => optimizer.Optimize(broken, PathOf("broken-out.xlsx")));
    }

    [Fact]
    public void Optimize_CancelledBeforeWritingLeavesNoPartialFilesOrOutput()
    {
        var input = CreateSample();
        var output = PathOf("cancelled.xlsx");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new ExcelStyleOptimizer().Optimize(input, output, null, cancellation.Token));

        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(directory, "*.partial"));
    }

    [Fact]
    public void Analyze_RejectsStyleIndexesOutsideTheStyleTable()
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(DefaultXf));
        var sheet = OoxmlWorkbookFixture.Worksheet("A1",
            """<row r="1"><c r="A1" s="99"><v>1</v></c></row>""");
        var path = OoxmlWorkbookFixture.Create(PathOf("invalid.xlsx"), styles, new SheetSpec("Sheet1", sheet));

        var error = Assert.Throws<InvalidDataException>(() => new ExcelStyleAnalyzer().Analyze(path));

        Assert.Contains("99", error.Message);
    }

    [Fact]
    public void Previews_ResolveThemeIndexedTintFormatsAndAlignment()
    {
        var fonts = """
            <fonts count="3">
              <font><sz val="11"/><color theme="4"/><name val="Calibri"/></font>
              <font><sz val="10"/><color indexed="2"/><name val="Calibri"/></font>
              <font><b/><sz val="12"/><color theme="4" tint="-0.5"/><name val="Calibri"/></font>
            </fonts>
            """;
        var fills = """
            <fills count="2">
              <fill><patternFill patternType="none"/></fill>
              <fill><patternFill patternType="solid"><fgColor indexed="5"/></patternFill></fill>
            </fills>
            """;
        var borders = """
            <borders count="2">
              <border><left/><right/><top/><bottom/><diagonal/></border>
              <border><left/><right/><top/><bottom style="medium"><color rgb="FF00FF00"/></bottom><diagonal/></border>
            </borders>
            """;
        var formats = """<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>""";
        var cellXfs = OoxmlWorkbookFixture.CellXfs(
            DefaultXf,
            """<xf numFmtId="164" fontId="1" fillId="1" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/>""",
            """<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>""");
        var sheet = OoxmlWorkbookFixture.Worksheet("A1:B1",
            """<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>""");
        var path = OoxmlWorkbookFixture.Create(PathOf("colors.xlsx"),
            OoxmlWorkbookFixture.StylesXml(cellXfs, fonts, fills, borders, formats), new SheetSpec("Sheet1", sheet));

        var analysis = new ExcelStyleAnalyzer().Analyze(path);
        var usage = analysis.Worksheets.Single().Styles;

        var formatsStyle = usage.Single(item => item.StyleId == 1).Style;
        Assert.Equal("#FF0000", formatsStyle.FontColor);
        Assert.Equal("#FFFF00", formatsStyle.FillColor);
        Assert.Equal("yyyy-mm-dd", formatsStyle.NumberFormat);
        Assert.Equal(BorderSides.Bottom, formatsStyle.Borders);
        Assert.Equal("#00FF00", formatsStyle.BorderColor);
        Assert.Equal(1.6, formatsStyle.BorderThickness);

        var alignedStyle = usage.Single(item => item.StyleId == 2).Style;
        Assert.True(alignedStyle.Bold);
        Assert.Equal(PreviewHorizontalAlignment.Center, alignedStyle.HorizontalAlignment);
        Assert.Equal(PreviewVerticalAlignment.Top, alignedStyle.VerticalAlignment);
        Assert.True(alignedStyle.WrapText);
        Assert.NotEqual("#4472C4", alignedStyle.FontColor);
        Assert.True(Luminance(alignedStyle.FontColor) < Luminance("#4472C4"));
    }

    [Fact]
    public void Analyze_ReportsPerWorksheetAndRowProgress()
    {
        var input = CreateSample();
        var progress = new ProgressCollector<ExcelAnalysisProgress>();

        var analysis = new ExcelStyleAnalyzer().Analyze(input, default, progress);

        Assert.Contains(progress.Reports, report => report.Phase == ExcelAnalysisPhase.ReadingStyles);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelAnalysisPhase.Projecting);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelAnalysisPhase.ScanningWorksheets &&
                                                    report.Worksheet == "Data" && report.RowsScanned >= 1);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelAnalysisPhase.ScanningWorksheets &&
                                                    report.Completed == analysis.SheetCount);
        Assert.All(progress.Reports, report => Assert.InRange(report.Completed, 0, analysis.SheetCount));
    }

    [Fact]
    public void Optimize_ReportsPerWorksheetRewritingAndVerification()
    {
        var input = CreateSample();
        var output = PathOf("progress.xlsx");
        var progress = new ProgressCollector<ExcelOptimizationProgress>();

        var result = new ExcelStyleOptimizer().Optimize(input, output, progress);

        Assert.True(result.Verified);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelOptimizationPhase.RewritingWorksheets && report.Completed == 2);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelOptimizationPhase.Verifying && report.Completed == 2);
        Assert.Contains(progress.Reports, report => report.Phase == ExcelOptimizationPhase.WritingOutput);
    }

    private static void AssertSameVisualStyles(ExcelStyleAnalysis before, ExcelStyleAnalysis after)
    {
        Assert.Equal(before.Worksheets.Count, after.Worksheets.Count);
        foreach (var sheetBefore in before.Worksheets)
        {
            var sheetAfter = after.Worksheets.Single(sheet => sheet.Name == sheetBefore.Name);
            Assert.Equal(sheetBefore.CellCount, sheetAfter.CellCount);
            Assert.Equal(sheetBefore.UsedRange, sheetAfter.UsedRange);
            var expected = DescribeUsage(sheetBefore);
            var actual = DescribeUsage(sheetAfter);
            Assert.Equal(expected.Count, actual.Count);
            foreach (var (key, count) in expected) Assert.Equal(count, actual.GetValueOrDefault(key));
        }
    }

    private static Dictionary<string, int> DescribeUsage(WorksheetStyleSummary sheet)
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var usage in sheet.Styles)
        {
            var key = Describe(usage.Style);
            result[key] = result.GetValueOrDefault(key) + usage.CellCount;
        }
        return result;
    }

    private static string Describe(CellStylePreview style) =>
        string.Join('|', style.IsDefault, style.FillColor, style.FillPattern, style.FontName,
            style.FontSize.ToString(CultureInfo.InvariantCulture), style.Bold, style.Italic, style.Underline,
            style.Strike, style.FontColor, style.Borders, style.BorderColor,
            style.BorderThickness.ToString(CultureInfo.InvariantCulture),
            style.HorizontalAlignment, style.VerticalAlignment, style.WrapText, style.NumberFormat);

    private static double Luminance(string hex) =>
        (0.2126 * Convert.ToInt32(hex.Substring(1, 2), 16)) +
        (0.7152 * Convert.ToInt32(hex.Substring(3, 2), 16)) +
        (0.0722 * Convert.ToInt32(hex.Substring(5, 2), 16));

    private sealed class ProgressCollector<T> : IProgress<T>
    {
        public List<T> Reports { get; } = new();
        public void Report(T value) => Reports.Add(value);
    }
}
