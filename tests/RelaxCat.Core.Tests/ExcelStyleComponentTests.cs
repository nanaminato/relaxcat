using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using RelaxCat.Core.Excel;
using Xunit;

namespace RelaxCat.Core.Tests;

public sealed class ExcelStyleComponentTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-components-" + Guid.NewGuid().ToString("N"));
    private static readonly XNamespace Ns = OoxmlWorkbookFixture.MainNamespace;
    private static readonly ExcelStyleOptimizationOptions Gc = new(CollectUnusedStyles: true);
    public ExcelStyleComponentTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);
    private string PathOf(string name) => Path.Combine(directory, name);
    private static string Xf(int font = 0, int fill = 0, int border = 0, int number = 0, int? baseId = null) =>
        $"<xf fontId='{font}' fillId='{fill}' borderId='{border}' numFmtId='{number}'" +
        (baseId is null ? "/>" : $" xfId='{baseId}'/>");

    private static string Styles()
    {
        var xml = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            Xf(baseId: 0), Xf(1, 2, 1, 164, 1), Xf(2, 3, 2, 165, 2), Xf(3, 4, 3, 166, 3)),
            "<fonts count='4'><font/><font><b/></font><font><b/></font><font><i/></font></fonts>",
            "<fills count='5'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill>" +
            "<fill><patternFill patternType='solid'><fgColor rgb='FFFF0000'/></patternFill></fill>" +
            "<fill><patternFill patternType='solid'><fgColor rgb='FFFF0000'/></patternFill></fill>" +
            "<fill><patternFill patternType='solid'><fgColor rgb='FF0000FF'/></patternFill></fill></fills>",
            "<borders count='4'><border/><border><left style='thin'/></border><border><left style='thin'/></border><border><right style='thick'/></border></borders>",
            "<numFmts count='3'><numFmt numFmtId='164' formatCode='0.000'/><numFmt numFmtId='165' formatCode='0.000'/><numFmt numFmtId='166' formatCode='0.0000'/></numFmts>");
        var document = XDocument.Parse(xml);
        document.Root!.Element(Ns + "cellStyleXfs")!.ReplaceWith(XElement.Parse(
            $"<cellStyleXfs xmlns='{Ns}' count='4'>" + Xf() + Xf(1, 2, 1, 164) + Xf(2, 3, 2, 165) + Xf(3, 4, 3, 166) + "</cellStyleXfs>"));
        document.Root.Element(Ns + "cellStyles")!.ReplaceWith(XElement.Parse(
            $"<cellStyles xmlns='{Ns}' count='2'><cellStyle name='Normal' xfId='0' builtinId='0'/><cellStyle name='Named' xfId='2'/></cellStyles>"));
        return document.ToString();
    }

    private string Create(string styles, string? sheet = null) => OoxmlWorkbookFixture.Create(PathOf("input.xlsx"), styles,
        new SheetSpec("Synthetic", sheet ?? OoxmlWorkbookFixture.Worksheet("A1:B2",
            "<row r='1' s='1'><c r='A1' s='2'/><c r='B1'/></row><row r='2'><c r='A2'/></row>")));
    private static XDocument ReadStyles(string path) => XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(path, "xl/styles.xml")));
    private static int Count(XDocument document, string section) => document.Root!.Element(Ns + section)!.Elements().Count();


    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16")]
    [InlineData("utf-16BE")]
    public void GcPreservesEncodingAndUntouchedPartBytes(string encodingName)
    {
        var encoding = Encoding.GetEncoding(encodingName);
        var originalXml = Styles();
        // XDocument fixture has no declaration: add one matching the requested encoding.
        originalXml = $"<?xml version='1.0' encoding='{encodingName}'?>" + originalXml;
        var input = Create(Styles());
        using (var archive = ZipFile.Open(input, ZipArchiveMode.Update)) archive.GetEntry("xl/styles.xml")!.Delete();
        OoxmlWorkbookFixture.AddPart(input, "xl/styles.xml", encoding.GetPreamble().Concat(encoding.GetBytes(originalXml)).ToArray());
        var source = File.ReadAllBytes(input);
        var output = PathOf("encoded.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output, options: Gc);
        Assert.True(result.Verified);
        Assert.Equal(source, File.ReadAllBytes(input));
        var actual = OoxmlWorkbookFixture.ReadPart(output, "xl/styles.xml");
        Assert.True(actual.AsSpan().StartsWith(encoding.GetPreamble()));
        Assert.Contains($"encoding='{encodingName}'", encoding.GetString(actual));
        foreach (var part in OoxmlWorkbookFixture.PartNames(input).Where(p => p != "xl/styles.xml" && !p.StartsWith("xl/worksheets/")))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(output, part));
        Assert.Equal(2, new ExcelStyleAnalyzer().Analyze(output, options: Gc).CellXfCount);
    }


    [Fact]
    public void SparseCustomNumberFormatIdsAreKeysRatherThanTablePositions()
    {
        var styles = Styles().Replace("numFmtId=\"164\"", "numFmtId=\"4096\"")
            .Replace("numFmtId=\"165\"", "numFmtId=\"65535\"").Replace("numFmtId=\"166\"", "numFmtId=\"8192\"");
        var input = Create(styles);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("sparse.xlsx"), options: Gc);
        Assert.True(result.Verified);
        var after = ReadStyles(result.OutputPath);
        Assert.Equal("4096", Assert.Single(after.Root!.Element(Ns + "numFmts")!.Elements()).Attribute("numFmtId")!.Value);
        Assert.Equal("4096", after.Root.Element(Ns + "cellXfs")!.Elements().Last().Attribute("numFmtId")!.Value);
    }

    [Fact]
    public void UnknownStyleReferenceLocationDisablesAdvancedChanges()
    {
        var document = XDocument.Parse(Styles());
        document.Root!.Element(Ns + "dxfs")!.ReplaceWith(XElement.Parse($"<dxfs xmlns='{Ns}' count='1'><dxf fontId='3'/></dxfs>"));
        var input = Create(document.ToString());
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("unknown-ref.xlsx"), options: Gc);
        Assert.Equal("unsupportedReferences", result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
    }

    [Fact]
    public void CanonicalKeysDoNotConfuseAttributeDelimitersWithDefinitions()
    {
        var document = XDocument.Parse(Styles());
        document.Root!.Element(Ns + "fonts")!.ReplaceWith(XElement.Parse(
            $"<fonts xmlns='{Ns}' count='4'><font/><font><name val='a;:x=y'/></font><font><name val='a' x='y'/></font><font><i/></font></fonts>"));
        var input = Create(document.ToString());
        var output = PathOf("delimiters.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output);
        Assert.True(result.Verified);
        Assert.Equal(4, Count(ReadStyles(output), "fonts"));
    }

    [Fact]
    public void LargeComponentTableRemapsHighIndicesWithoutChangingStyleMeaning()
    {
        const int count = 20000;
        var styles = XDocument.Parse(Styles());
        styles.Root!.Element(Ns + "fonts")!.ReplaceWith(XElement.Parse(
            $"<fonts xmlns='{Ns}' count='{count}'>" + "<font/>" + string.Concat(Enumerable.Repeat("<font><b/></font>", count - 1)) + "</fonts>"));
        foreach (var xf in styles.Descendants(Ns + "xf"))
            if ((string?)xf.Attribute("fontId") != "0") xf.SetAttributeValue("fontId", count - 1);
        var input = Create(styles.ToString());
        var output = PathOf("large.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output);
        Assert.True(result.Verified);
        Assert.Equal(count - 2, result.TableChanges.Single(c => c.Name == "fonts").DuplicatesRemoved);
        Assert.Equal(2, Count(ReadStyles(output), "fonts"));
    }

    [Fact]
    public void ReservedFillSlotsAndDifferentApplyFlagsAreNeverMergedAway()
    {
        var styles = XDocument.Parse(Styles());
        var fills = styles.Root!.Element(Ns + "fills")!;
        fills.Elements().ElementAt(1).ReplaceWith(new XElement(fills.Elements().First()));
        var formats = styles.Root.Element(Ns + "cellXfs")!.Elements().ToArray();
        formats[1].SetAttributeValue("applyFont", 0);
        formats[2].SetAttributeValue("applyFont", 1);
        var input = Create(styles.ToString());
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("reserved.xlsx"), options: Gc);
        Assert.True(result.Verified);
        var after = ReadStyles(result.OutputPath);
        Assert.Equal(3, Count(after, "cellXfs"));
        Assert.Equal(3, Count(after, "fills"));
        Assert.Equal(after.Root!.Element(Ns + "fills")!.Elements().First().ToString(),
            after.Root.Element(Ns + "fills")!.Elements().ElementAt(1).ToString());
    }

    [Fact]
    public void DuplicateCustomNumberFormatIdsAreRejectedBeforePublishing()
    {
        var styles = Styles().Replace("numFmtId=\"165\" formatCode=\"0.000\"", "numFmtId=\"164\" formatCode=\"0.000\"");
        var input = Create(styles);
        var output = PathOf("invalid.xlsx");
        Assert.Throws<InvalidDataException>(() => new ExcelStyleOptimizer().Optimize(input, output, options: Gc));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void ComponentsMergeThenUnlockBaseAndCellFormatMerging()
    {
        var input = Create(Styles());
        var before = File.ReadAllBytes(input);
        var analysis = new ExcelStyleAnalyzer().Analyze(input);
        Assert.Null(analysis.AdvancedOptimizationSkipReason);
        Assert.True(analysis.HasStyleChanges);
        var output = PathOf("output.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output);
        Assert.True(result.Verified);
        Assert.Equal(analysis.ProjectedStylesXmlBytes, result.OutputStylesXmlBytes);
        var after = ReadStyles(output);
        Assert.Equal(3, Count(after, "fonts"));
        Assert.Equal(4, Count(after, "fills"));
        Assert.Equal(3, Count(after, "borders"));
        Assert.Equal(2, Count(after, "numFmts"));
        Assert.Equal(3, Count(after, "cellStyleXfs"));
        Assert.Equal(3, Count(after, "cellXfs"));
        Assert.All(result.TableChanges, change => Assert.Equal(0, change.UnusedRemoved));
        Assert.Equal(before, File.ReadAllBytes(input));
    }

    [Fact]
    public void GcCollectsUnreachableChainAndPreservesDefaultsNamedStylesAndAppearance()
    {
        var input = Create(Styles());
        var analyzer = new ExcelStyleAnalyzer();
        var before = analyzer.Analyze(input, options: Gc);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("gc.xlsx"), options: Gc);
        Assert.True(result.Verified);
        Assert.Equal(before.ProjectedStylesXmlBytes, result.OutputStylesXmlBytes);
        var after = analyzer.Analyze(result.OutputPath, options: Gc);
        Assert.False(after.HasStyleChanges);
        var document = ReadStyles(result.OutputPath);
        foreach (var section in new[] { "fonts", "borders", "cellStyleXfs", "cellXfs" }) Assert.Equal(2, Count(document, section));
        Assert.Equal(3, Count(document, "fills"));
        Assert.Equal(1, Count(document, "numFmts"));
        Assert.Equal(2, Count(document, "cellStyles"));
        Assert.Equal("1", document.Root!.Element(Ns + "cellStyles")!.Elements().Last().Attribute("xfId")!.Value);
        Assert.All(result.TableChanges, change => Assert.Equal(change.Before - change.After, change.DuplicatesRemoved + change.UnusedRemoved));
        Assert.True(result.TableChanges.Sum(change => change.UnusedRemoved) >= 6);
        var styleBefore = before.Worksheets[0].Styles.Single(s => s.StyleId == 2).Style;
        var styleAfter = after.Worksheets[0].Styles.Single(s => s.Style.Bold).Style;
        Assert.Equal(styleBefore with { StyleId = styleAfter.StyleId }, styleAfter);
        var second = PathOf("second.xlsx");
        new ExcelStyleOptimizer().Optimize(result.OutputPath, second, options: Gc);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(second, "xl/styles.xml"));
    }

    [Fact]
    public void NamedStyleAndColumnReferencesKeepOtherwiseUnusedDefinitionsAlive()
    {
        var document = XDocument.Parse(Styles());
        document.Root!.Element(Ns + "cellStyles")!.Elements().Last().SetAttributeValue("xfId", 3);
        var sheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'/></row>")
            .Replace("<sheetData>", "<cols><col min='4' max='8' style='3'/></cols><sheetData>");
        var input = Create(document.ToString(), sheet);
        var output = PathOf("roots.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output, options: Gc);
        var after = ReadStyles(output);
        Assert.True(result.Verified);
        Assert.Equal(2, Count(after, "cellXfs"));
        Assert.Equal(2, Count(after, "cellStyleXfs"));
        Assert.Equal(2, Count(after, "fonts")); // italic font survives through named and column roots
        Assert.Equal(1, Count(after, "numFmts"));
        var outputSheet = XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(output, "xl/worksheets/sheet1.xml")));
        Assert.Equal("1", outputSheet.Descendants(Ns + "col").Single().Attribute("style")!.Value);
    }

    [Theory]
    [InlineData("xl/pivotTables/pivotTable1.xml")]
    [InlineData("xl/tables/table1.xml")]
    [InlineData("xl/vbaProject.bin")]
    [InlineData("customXml/item1.xml")]
    public void UnknownPackageDependenciesDisableAdvancedEditsAndPreserveParts(string part)
    {
        var input = Create(Styles());
        byte[] payload = [11, 22, 33];
        OoxmlWorkbookFixture.AddPart(input, part, payload);
        var analysis = new ExcelStyleAnalyzer().Analyze(input, options: Gc);
        Assert.Equal("unsupportedParts", analysis.AdvancedOptimizationSkipReason);
        var output = PathOf("safe.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output, options: Gc);
        Assert.Equal("unsupportedParts", result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(output, "xl/styles.xml"));
        Assert.Equal(payload, OoxmlWorkbookFixture.ReadPart(output, part));
    }

    [Fact]
    public void WorksheetExtensionDisablesAdvancedEdits()
    {
        var sheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'/></row>")
            .Replace("</worksheet>", "<extLst><ext uri='synthetic'><hidden xmlns='urn:unknown' fontId='3'/></ext></extLst></worksheet>");
        var input = Create(Styles(), sheet);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("extension.xlsx"), options: Gc);
        Assert.Equal("packageExtensions", result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
    }

    [Fact]
    public void DifferentialNumberFormatsRetainTheirIndependentIdsAndGlobalTable()
    {
        var styles = XDocument.Parse(Styles());
        styles.Root!.Element(Ns + "dxfs")!.ReplaceWith(XElement.Parse(
            $"<dxfs xmlns='{Ns}' count='1'><dxf><numFmt numFmtId='166' formatCode='0.0%'/><font><i/></font></dxf></dxfs>"));
        var input = Create(styles.ToString());
        var output = PathOf("dxf.xlsx");
        Assert.True(new ExcelStyleOptimizer().Optimize(input, output, options: Gc).Verified);
        var after = ReadStyles(output);
        Assert.Equal(3, Count(after, "numFmts"));
        Assert.Equal(styles.Root.Element(Ns + "dxfs")!.ToString(), after.Root!.Element(Ns + "dxfs")!.ToString());
    }

    [Fact]
    public void ComponentsCanBeDisabledAndGcRemainsOptIn()
    {
        var input = Create(Styles());
        var output = PathOf("disabled.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output, options: new(DeduplicateComponents: false));
        Assert.True(result.Verified);
        var before = ReadStyles(input);
        var after = ReadStyles(output);
        foreach (var section in new[] { "fonts", "fills", "borders", "numFmts" })
            Assert.Equal(before.Root!.Element(Ns + section)!.ToString(), after.Root!.Element(Ns + section)!.ToString());
        // Semantic XF equality can be proven through equal component definitions
        // without physically deduplicating the component tables themselves.
        Assert.Equal(3, Count(after, "cellXfs"));
        Assert.Equal(3, Count(after, "cellStyleXfs"));
        Assert.All(result.TableChanges, change => Assert.Equal(0, change.UnusedRemoved));
    }

    [Fact]
    public void EmptyWorkbookRetainsImplicitFormatAndReservedFills()
    {
        var input = OoxmlWorkbookFixture.Create(PathOf("input.xlsx"), Styles());
        var output = PathOf("empty.xlsx");
        Assert.True(new ExcelStyleOptimizer().Optimize(input, output, options: Gc).Verified);
        var styles = ReadStyles(output);
        Assert.Equal(1, Count(styles, "cellXfs"));
        Assert.Equal(3, Count(styles, "fills")); // reserved fills plus the named style's red fill
        Assert.Equal(2, Count(styles, "cellStyleXfs")); // named style remains a root
    }

    [Fact]
    public void BatchForwardsGcOptions()
    {
        var input = Create(Styles());
        var outputDirectory = PathOf("batch");
        var outcome = new ExcelBatchOptimizer().Optimize([new ExcelBatchFile(input, "input.xlsx", new FileInfo(input).Length)], outputDirectory, false, options: Gc);
        var item = Assert.Single(outcome.Results);
        Assert.Equal(ExcelBatchItemStatus.Completed, item.Status);
        Assert.Equal(2, Count(ReadStyles(item.OutputPath!), "cellXfs"));
    }
}
