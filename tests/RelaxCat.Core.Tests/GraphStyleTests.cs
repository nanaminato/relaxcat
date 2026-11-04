using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using RelaxCat.Core.Excel;
using Xunit;
using Xunit.Abstractions;

namespace RelaxCat.Core.Tests;

/// <summary>Generated adversarial packages only; no private workbook is needed.</summary>
public sealed class GraphStyleTests : IDisposable
{
    private static readonly XNamespace Ns = OoxmlWorkbookFixture.MainNamespace;
    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-style-graph-" + Guid.NewGuid().ToString("N"));
    private readonly ITestOutputHelper output;
    public GraphStyleTests(ITestOutputHelper output)
    {
        this.output = output;
        Directory.CreateDirectory(directory);
    }
    public void Dispose() => Directory.Delete(directory, true);
    private string PathOf(string name) => Path.Combine(directory, name);
    private static string Xf(int baseId = 0, int fontId = 0) =>
        $"<xf numFmtId='0' fontId='{fontId}' fillId='0' borderId='0' xfId='{baseId}'/>";
    private static string Base(string extra = "", int fontId = 0) =>
        $"<xf numFmtId='0' fontId='{fontId}' fillId='0' borderId='0' {extra}/>";

    private static string Styles(IEnumerable<string> bases, IEnumerable<string> cells, IEnumerable<string>? named = null)
    {
        var document = XDocument.Parse(OoxmlWorkbookFixture.StylesXml(
            OoxmlWorkbookFixture.CellXfs(cells.ToArray()),
            "<fonts count='2'><font><name val='Calibri'/></font><font><name val='Calibri'/><b/></font></fonts>"));
        var baseArray = bases.ToArray();
        document.Root!.Element(Ns + "cellStyleXfs")!.ReplaceWith(XElement.Parse(
            $"<cellStyleXfs xmlns='{Ns}' count='{baseArray.Length}'>" + string.Concat(baseArray) + "</cellStyleXfs>"));
        if (named is not null)
        {
            var names = named.ToArray();
            document.Root.Element(Ns + "cellStyles")!.ReplaceWith(XElement.Parse(
                $"<cellStyles xmlns='{Ns}' count='{names.Length}'>" + string.Concat(names) + "</cellStyles>"));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private string Create(string styles, string? worksheet = null) => OoxmlWorkbookFixture.Create(PathOf("input.xlsx"), styles,
        new SheetSpec("Synthetic", worksheet ?? OoxmlWorkbookFixture.Worksheet("A1",
            "<row r='1'><c r='A1' s='0'><v>7</v></c></row>")));
    private static XDocument Part(string path, string part = "xl/styles.xml") =>
        XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(path, part)));
    private static XElement[] Table(string path, string table) => Part(path).Root!.Element(Ns + table)!.Elements().ToArray();

    [Fact]
    public void FortyEightThousandNamedStylesRetainDistinctIdentitiesAndWorkbookMeaning()
    {
        const int count = 48846;
        var bases = Enumerable.Range(0, count).Select(i => i == 0 ? Base() : Base(i % 2 == 0 ? "applyFont='true'" : "applyFont='1'", 1));
        var names = Enumerable.Range(0, count).Select(i =>
            $"<cellStyle name='{(i == 0 ? "Normal" : "Synthetic_" + i)}' xfId='{i}'{(i == 0 ? " builtinId='0'" : "")}/>");
        var sheet = $"""
            <worksheet xmlns="{Ns}">
              <dimension ref="A1:D3"/>
              <cols><col min="4" max="4" width="22" customWidth="1" style="2"/></cols>
              <sheetData>
                <row r="1"><c r="A1" s="1"><f>SUM(B1:C1)</f><v>5</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"/></row>
                <row r="2" s="2" customFormat="1"><c r="A2" t="inlineStr"><is><t>preserve me</t></is></c></row>
                <row r="3"><c r="A3" s="2"/></row>
              </sheetData>
              <mergeCells count="1"><mergeCell ref="A3:C3"/></mergeCells>
            </worksheet>
            """;
        var input = Create(Styles(bases, [Xf(), Xf(1, 1), Xf(count - 1, 1)], names), sheet);
        var original = File.ReadAllBytes(input);
        var analyzer = new ExcelStyleAnalyzer();
        var before = analyzer.Analyze(input);
        Assert.True(before.ReferenceGraphComplete);
        var baseDiagnostic = before.Sections.Single(s => s.Name == "cellStyleXfs");
        Assert.Equal(count - 3, baseDiagnostic.DuplicateCount);
        Assert.Equal(count - 2, baseDiagnostic.SemanticDuplicateCount);
        Assert.Equal(0, baseDiagnostic.SafeDuplicateRemovalCount);
        Assert.Equal(0, baseDiagnostic.UnreferencedCount);
        Assert.Contains(before.ReferenceGraph, edge => edge.Source == "cellStyles" &&
            edge.SourceIndex == count - 1 && edge.Target == "cellStyleXfs" && edge.TargetIndex == count - 1);
        Assert.Contains(before.ReferenceGraph, edge => edge.Source == "cellXfs" &&
            edge.SourceIndex == 2 && edge.Target == "cellStyleXfs" && edge.TargetIndex == count - 1);
        Assert.Contains(before.ReferenceGraph, edge => edge.Source == "xl/worksheets/sheet1.xml" &&
            edge.Target == "cellXfs" && edge.TargetIndex == 2);
        var output = PathOf("optimized.xlsx");
        var result = new ExcelStyleOptimizer().Optimize(input, output);

        Assert.True(result.Verified);
        Assert.All(result.Checks, check => Assert.True(check.Passed, check.Detail));
        Assert.Equal(count, result.OutputCellStyleXfCount);
        Assert.Equal(3, result.OutputCellXfCount);
        Assert.Equal(new[] { 0, 1, 2 }, result.CellXfMap);
        Assert.Equal(count, result.CellStyleXfMap.Count);
        Assert.Equal(0, result.CellStyleXfMap[0]);
        Assert.Equal(Enumerable.Range(0, count), result.CellStyleXfMap);
        Assert.Equal(before.ProjectedStylesXmlBytes, result.OutputStylesXmlBytes);
        Assert.Equal(result.InputStylesXmlBytes, result.OutputStylesXmlBytes);
        Assert.Equal(original, File.ReadAllBytes(input));
        Assert.Equal(count, Table(output, "cellStyles").Length);
        var oldNames = Table(input, "cellStyles");
        var newNames = Table(output, "cellStyles");
        for (var i = 0; i < count; i++)
        {
            Assert.Equal(oldNames[i].Attribute("name")!.Value, newNames[i].Attribute("name")!.Value);
            Assert.Equal(i.ToString(), newNames[i].Attribute("xfId")!.Value);
        }
        Assert.Equal(new[] { "0", "1", (count - 1).ToString() }, Table(output, "cellXfs").Select(x => x.Attribute("xfId")!.Value));
        Assert.Equal(WorksheetWithoutStyleReferences(input), WorksheetWithoutStyleReferences(output));
        var mappedSheet = Part(output, "xl/worksheets/sheet1.xml");
        Assert.Equal("2", mappedSheet.Descendants(Ns + "col").Single().Attribute("style")!.Value);
        Assert.Equal("2", mappedSheet.Descendants(Ns + "row").Single(x => (string?)x.Attribute("r") == "2").Attribute("s")!.Value);
        var after = analyzer.Analyze(output);
        Assert.Equal(Usage(before), Usage(after));
        foreach (var part in OoxmlWorkbookFixture.PartNames(input).Where(p => p != "xl/styles.xml" && p != "xl/worksheets/sheet1.xml"))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(output, part));

        // Measure uncompressed styles XML, compressed entry and whole package independently.
        using (var archive = ZipFile.OpenRead(input))
        {
            var entry = archive.GetEntry("xl/styles.xml")!;
            Assert.Equal(entry.Length, before.StylesXmlBytes);
            Assert.Equal(entry.CompressedLength, before.StylesXmlCompressedBytes);
            Assert.True(entry.CompressedLength < entry.Length);
        }
        Assert.Equal(new FileInfo(input).Length, result.InputFileSize);
        Assert.Equal(new FileInfo(output).Length, result.OutputFileSize);
        this.output.WriteLine($"Synthetic {count:N0} base/named styles: XML {result.InputStylesXmlBytes:N0} -> {result.OutputStylesXmlBytes:N0} bytes; ZIP style entry {before.StylesXmlCompressedBytes:N0} -> {after.StylesXmlCompressedBytes:N0} bytes; workbook {result.InputFileSize:N0} -> {result.OutputFileSize:N0} bytes.");
        var second = PathOf("second.xlsx");
        var rerun = new ExcelStyleOptimizer().Optimize(output, second);
        Assert.True(rerun.Verified);
        Assert.Equal(0, rerun.RemappedReferences);
        foreach (var part in OoxmlWorkbookFixture.PartNames(output))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(output, part), OoxmlWorkbookFixture.ReadPart(second, part));
    }

    [Fact]
    public void SemanticBooleanEqualityDoesNotEraseMissingFlagsOrAlignmentProtectionDifferences()
    {
        var bases = new[]
        {
            Base(), Base("applyFont='false'"), Base("applyFont='0'"), Base("applyFont='true'"), Base("applyFont='1'"),
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'><alignment horizontal='left'/></xf>",
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'><alignment horizontal='right'/></xf>",
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'><protection locked='0'/></xf>",
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'><protection locked='1'/></xf>"
        };
        var input = Create(Styles(bases, Enumerable.Range(0, bases.Length).Select(i => Xf(i))));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("flags.xlsx"));
        Assert.True(result.Verified);
        Assert.Equal(7, result.OutputCellStyleXfCount);
        Assert.Equal(7, result.OutputCellXfCount);
        var outputBases = Table(result.OutputPath, "cellStyleXfs");
        Assert.Null(outputBases[0].Attribute("applyFont"));
        Assert.Equal(2, outputBases.Count(x => x.Element(Ns + "alignment") is not null));
        Assert.Equal(2, outputBases.Count(x => x.Element(Ns + "protection") is not null));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnknownReferenceLocationPreservesAllIndicesIncludingCellFormats(bool collectUnused)
    {
        var sheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' s='1'/></row>")
            .Replace("</worksheet>", "<extLst><ext uri='synthetic'><opaque xmlns='urn:unknown' xfId='1' s='1'/></ext></extLst></worksheet>");
        var input = Create(Styles([Base(), Base()], [Xf(), Xf()], ["<cellStyle name='Custom' xfId='0'/>", "<cellStyle name='Custom_1' xfId='1'/>"]), sheet);
        var analysis = new ExcelStyleAnalyzer().Analyze(input, options: new(CollectUnusedStyles: collectUnused, RemoveCopiedNamedStyles: true));
        Assert.False(analysis.ReferenceGraphComplete);
        Assert.All(analysis.Sections, section => Assert.Equal(0, section.SafeDuplicateRemovalCount));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("unknown.xlsx"),
            options: new(CollectUnusedStyles: collectUnused, RemoveCopiedNamedStyles: true));
        Assert.True(result.Verified);
        Assert.NotNull(result.AdvancedOptimizationSkipReason);
        Assert.Equal(2, result.OutputCellXfCount);
        Assert.Equal(2, result.OutputCellStyleXfCount);
        Assert.Empty(result.TableChanges);
        foreach (var part in OoxmlWorkbookFixture.PartNames(input))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
    }

    [Fact]
    public void MacroPackagePreservesAllIndicesAndBinaryPayload()
    {
        var input = OoxmlWorkbookFixture.Create(PathOf("macro.xlsm"),
            Styles([Base(), Base()], [Xf(), Xf(1)], ["<cellStyle name='Custom' xfId='0'/>", "<cellStyle name='Custom_1' xfId='1'/>"]),
            new SheetSpec("Synthetic", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' s='1'/></row>")));
        OoxmlWorkbookFixture.AddPart(input, "xl/vbaProject.bin", [1, 3, 5, 7]);
        var original = File.ReadAllBytes(input);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("macro-copy.xlsm"),
            options: new(CollectUnusedStyles: true, RemoveCopiedNamedStyles: true));
        Assert.True(result.Verified);
        Assert.Equal(new[] { 0, 1 }, result.NamedStyleMap);
        Assert.Equal("unsupportedParts", result.AdvancedOptimizationSkipReason);
        Assert.Equal(new[] { 0, 1 }, result.CellXfMap);
        Assert.Equal(new[] { 0, 1 }, result.CellStyleXfMap);
        Assert.Equal(original, File.ReadAllBytes(input));
        foreach (var part in OoxmlWorkbookFixture.PartNames(input))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
    }

    [Fact]
    public void NamedStyleTextContainingAnAttributeLookalikeIsNeverRewritten()
    {
        const string name = "A xfId='1'";
        var input = Create(Styles([Base(), Base("applyFont='0'"), Base("applyFont='0'"), Base()], [Xf(), Xf(3)],
            ["<cellStyle name='Normal' xfId='0' builtinId='0'/>",
             "<cellStyle name=\"A xfId='1'\" xfId='3'/>"]));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("quoted-name.xlsx"));
        Assert.True(result.Verified);
        Assert.Equal(3, result.OutputCellStyleXfCount);
        var named = Table(result.OutputPath, "cellStyles")[1];
        Assert.Equal(name, named.Attribute("name")!.Value);
        Assert.Equal("2", named.Attribute("xfId")!.Value);
    }

    [Fact]
    public void GarbageCollectionIsOptInAndPreservesNamedStyleOnlyReachability()
    {
        var input = Create(Styles([Base(), Base("applyFont='1'", 1), Base("applyFont='0'", 1)],
            [Xf(), Xf(1, 1), Xf(2, 1)],
            ["<cellStyle name='Normal' xfId='0' builtinId='0'/>", "<cellStyle name='NamedOnly' xfId='1'/>"]));
        var safe = new ExcelStyleOptimizer().Optimize(input, PathOf("safe.xlsx"));
        Assert.Equal(3, safe.OutputCellXfCount);
        Assert.Equal(3, safe.OutputCellStyleXfCount);
        var gc = new ExcelStyleOptimizer().Optimize(input, PathOf("gc.xlsx"), options: new(CollectUnusedStyles: true));
        Assert.True(gc.Verified);
        Assert.Equal(1, gc.OutputCellXfCount);
        Assert.Equal(2, gc.OutputCellStyleXfCount);
        Assert.Equal(2, Table(gc.OutputPath, "fonts").Length);
        Assert.Equal("NamedOnly", Table(gc.OutputPath, "cellStyles")[1].Attribute("name")!.Value);
        Assert.Equal("1", Table(gc.OutputPath, "cellStyles")[1].Attribute("xfId")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CopiedNamesRequireOptInAndPreserveMetadataAmbiguityAndDifferentFormats(bool gc)
    {
        var names = new[]
        {
            "<cellStyle name='Normal' xfId='0' builtinId='0'/>",
            "<cellStyle name='Custom' xfId='1'/>",
            "<cellStyle name='Custom_1' xfId='2'/>",
            "<cellStyle name='Custom_2' xfId='0'/>",
            "<cellStyle name='Custom_3' xfId='1' builtinId='1'/>",
            "<cellStyle name='Custom_4' xfId='1' customBuiltin='1'/>",
            "<cellStyle name='Custom_5' xfId='1' hidden='1'/>",
            "<cellStyle name='Custom_6' xfId='1' iLevel='1'/>",
            "<cellStyle name='Custom_1_1' xfId='1'/>",
            "<cellStyle name='Other_1' xfId='1'/>",
            "<cellStyle name='Custom_01' xfId='1'/>",
            "<cellStyle name='Ambiguous' xfId='1'/>",
            "<cellStyle name='Distinct' xfId='1'/>",
            "<cellStyle name='Ambiguous_1' xfId='1'/>",
            "<cellStyle name='Custom_7' xfId='1'/>",
            "<cellStyle name='Unrelated' xfId='1'/>"
        };
        // Each named style owns a distinct base, even when formatting is equal.
        names = names.Select((name, i) =>
        {
            var element = XElement.Parse(name);
            element.SetAttributeValue("xfId", i);
            return element.ToString(SaveOptions.DisableFormatting);
        }).ToArray();
        var bases = Enumerable.Range(0, names.Length).Select(i => i is 0 or 3 ? Base() : Base("applyFont='1'", 1)).ToArray();
        // Additional copied-looking styles remain live through cell-format associations.
        var input = Create(Styles(bases, [Xf(), Xf(13, 1), Xf(14, 1)], names));
        var safe = new ExcelStyleOptimizer().Optimize(input, PathOf("names-safe.xlsx"));
        Assert.Equal(names.Length, Table(safe.OutputPath, "cellStyles").Length);
        Assert.Equal(Enumerable.Range(0, names.Length), safe.NamedStyleMap);
        var options = new ExcelStyleOptimizationOptions(CollectUnusedStyles: gc, RemoveCopiedNamedStyles: true);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("names-clean.xlsx"), options: options);
        Assert.True(result.Verified);
        var removed = gc ? new HashSet<int> { 2, 13, 14 } : new HashSet<int> { 2 };
        Assert.Equal(names.Length - removed.Count, Table(result.OutputPath, "cellStyles").Length);
        Assert.Equal(Enumerable.Range(0, names.Length).Select(i => removed.Contains(i) ? -1 : i - removed.Count(r => r < i)), result.NamedStyleMap);
        Assert.Equal(Table(input, "cellStyles").Where((_, i) => !removed.Contains(i)).Select(x => x.Attribute("name")!.Value),
            Table(result.OutputPath, "cellStyles").Select(x => x.Attribute("name")!.Value));
        Assert.Equal(gc ? 1 : 3, result.OutputCellXfCount);
        Assert.Equal(names.Length - (gc ? 3 : 0), result.OutputCellStyleXfCount);
        var second = new ExcelStyleOptimizer().Optimize(result.OutputPath, PathOf("names-again.xlsx"), options: options);
        Assert.True(second.Verified);
        foreach (var part in OoxmlWorkbookFixture.PartNames(result.OutputPath))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(result.OutputPath, part), OoxmlWorkbookFixture.ReadPart(second.OutputPath, part));
    }

    [Theory]
    [InlineData("c", false)]
    [InlineData("row", false)]
    [InlineData("col", false)]
    [InlineData("c", true)]
    [InlineData("row", true)]
    [InlineData("col", true)]
    public void MalformedExplicitStyleReferencesAreRejectedEvenWithoutChanges(string element, bool blocked)
    {
        var body = element switch
        {
            "c" => "<row r='1'><c r='A1' s='oops'/></row>",
            "row" => "<row r='1' s='oops'><c r='A1'/></row>",
            _ => "<row r='1'><c r='A1'/></row>"
        };
        var sheet = OoxmlWorkbookFixture.Worksheet("A1", body);
        if (element == "col") sheet = sheet.Replace("<sheetData>", "<cols><col min='1' max='1' style='4294967295'/></cols><sheetData>");
        if (blocked) sheet = sheet.Replace("</worksheet>", "<extLst><ext uri='synthetic'><opaque xmlns='urn:unknown'/></ext></extLst></worksheet>");
        var input = Create(Styles([Base()], [Xf()]), sheet);
        var output = PathOf("invalid.xlsx");
        Assert.Throws<InvalidDataException>(() => new ExcelStyleAnalyzer().Analyze(input));
        Assert.Throws<InvalidDataException>(() => new ExcelStyleOptimizer().Optimize(input, output));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void FalseAndZeroApplyAliasesPreserveContrastingBaseFontAndAlignment()
    {
        var baseStyle = "<xf numFmtId='0' fontId='1' fillId='0' borderId='0'><alignment horizontal='right'/></xf>";
        string Cell(string flag) => $"<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='1' applyFont='{flag}' applyAlignment='{flag}'><alignment horizontal='left'/></xf>";
        var sheet = OoxmlWorkbookFixture.Worksheet("A1:B1", "<row r='1'><c r='A1' s='1'/><c r='B1' s='2'/></row>");
        var input = Create(Styles([Base(), baseStyle], [Xf(), Cell("false"), Cell("0")]), sheet);
        var before = new ExcelStyleAnalyzer().Analyze(input);
        Assert.All(before.Worksheets.Single().Styles, usage =>
        {
            Assert.True(usage.Style.Bold);
            Assert.Equal(PreviewHorizontalAlignment.Right, usage.Style.HorizontalAlignment);
        });
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("inheritance.xlsx"));
        Assert.True(result.Verified);
        Assert.Equal(2, result.OutputCellXfCount);
        Assert.Equal(Usage(before), Usage(new ExcelStyleAnalyzer().Analyze(result.OutputPath)));
    }

    [Fact]
    public void TwentyThousandCellFormatsMapEveryOriginalIndex()
    {
        const int count = 20000;
        var input = Create(Styles([Base()], Enumerable.Repeat(Xf(), count)),
            OoxmlWorkbookFixture.Worksheet("A1", $"<row r='1'><c r='A1' s='{count - 1}'/></row>"));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("many-cells.xlsx"));
        Assert.True(result.Verified);
        Assert.Equal(1, result.OutputCellXfCount);
        Assert.Equal(count, result.CellXfMap.Count);
        Assert.All(result.CellXfMap, index => Assert.Equal(0, index));
        Assert.Equal("0", Part(result.OutputPath, "xl/worksheets/sheet1.xml").Descendants(Ns + "c").Single().Attribute("s")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EqualLookingNamedStylesKeepIndependentCellAssociationsWhenEdited(bool gc)
    {
        var input = Create(Styles([Base(), Base(), Base()], [Xf(), Xf(1).Replace("/>", " applyFont='0'/>"), Xf(2).Replace("/>", " applyFont='0'/>")],
            ["<cellStyle name='Normal' xfId='0' builtinId='0'/>", "<cellStyle name='Custom' xfId='1'/>", "<cellStyle name='Custom_1' xfId='2'/>"]),
            OoxmlWorkbookFixture.Worksheet("A1:B1", "<row r='1'><c r='A1' s='1'/><c r='B1' s='2'/></row>"));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("identity.xlsx"),
            options: new(CollectUnusedStyles: gc, RemoveCopiedNamedStyles: true));
        Assert.True(result.Verified);
        Assert.Equal(3, result.OutputCellStyleXfCount);
        Assert.Equal(3, result.OutputCellXfCount);
        Assert.Equal(new[] { 0, 1, 2 }, result.NamedStyleMap);
        var styles = Part(result.OutputPath);
        var named = styles.Root!.Element(Ns + "cellStyles")!.Elements().ToArray();
        Assert.Equal(3, named.Select(n => n.Attribute("xfId")!.Value).Distinct().Count());
        var bases = styles.Root.Element(Ns + "cellStyleXfs")!.Elements().ToArray();
        var customId = int.Parse(named.Single(n => (string?)n.Attribute("name") == "Custom").Attribute("xfId")!.Value);
        // Emulate editing Custom in Excel: only its own associated cell should inherit bold.
        var fonts = styles.Root.Element(Ns + "fonts")!;
        var newFontId = fonts.Elements().Count();
        fonts.Add(new XElement(Ns + "font", new XElement(Ns + "b")));
        fonts.SetAttributeValue("count", newFontId + 1);
        bases[customId].SetAttributeValue("fontId", newFontId);
        var edited = OoxmlWorkbookFixture.Create(PathOf("edited.xlsx"), styles.ToString(),
            new SheetSpec("Synthetic", Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/worksheets/sheet1.xml"))));
        var preview = new ExcelStyleAnalyzer().Analyze(edited).Worksheets.Single().Styles;
        Assert.True(preview.Single(s => s.StyleId == 1).Style.Bold);
        Assert.False(preview.Single(s => s.StyleId == 2).Style.Bold);
    }

    [Theory]
    [InlineData("<cellStyle name='Other' xfId='0'/>")]
    [InlineData("<cellStyle name='Other' xfId='00'/>")]
    [InlineData("<cellStyle name='Other'/>")]
    [InlineData("<cellStyle name='NORMAL' xfId='1'/>")]
    [InlineData("<cellStyle name='Other' xfId='oops'/>")]
    public void InvalidNamedIdentityIsRejectedBeforePublication(string invalidNamed)
    {
        var input = Create(Styles([Base(), Base()], [Xf()],
            ["<cellStyle name='Normal' xfId='0' builtinId='0'/>", invalidNamed]));
        var target = PathOf("invalid-named.xlsx");
        Assert.Throws<InvalidDataException>(() => new ExcelStyleAnalyzer().Analyze(input));
        Assert.Throws<InvalidDataException>(() => new ExcelStyleOptimizer().Optimize(input, target,
            options: new(CollectUnusedStyles: true, RemoveCopiedNamedStyles: true)));
        Assert.False(File.Exists(target));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedBaseAfterEquivalentUnownedBaseRetainsIdentityAcrossGcIndexShift(bool gc)
    {
        var input = Create(Styles([Base(), Base("applyFont='1'", 1), Base("applyFont='1'", 1)],
            [Xf(), Xf(2, 1)],
            ["<cellStyle name='Normal' xfId='0' builtinId='0'/>", "<cellStyle name='Custom' xfId='2'/>"]),
            OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' s='1'/></row>"));
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("named-order.xlsx"),
            options: new(CollectUnusedStyles: gc));
        Assert.True(result.Verified);
        Assert.Equal(gc ? 2 : 3, result.OutputCellStyleXfCount);
        Assert.Equal(gc ? new[] { 0, -1, 1 } : new[] { 0, 1, 2 }, result.CellStyleXfMap);
        Assert.Equal(new[] { 0, 1 }, result.NamedStyleMap);
        var named = Table(result.OutputPath, "cellStyles");
        var expectedBase = gc ? "1" : "2";
        Assert.Equal(expectedBase, named.Single(n => (string?)n.Attribute("name") == "Custom").Attribute("xfId")!.Value);
        Assert.Equal(expectedBase, Table(result.OutputPath, "cellXfs")[1].Attribute("xfId")!.Value);
        Assert.Equal(2, named.Select(n => n.Attribute("xfId")!.Value).Distinct().Count());
        Assert.Equal(Usage(new ExcelStyleAnalyzer().Analyze(input)), Usage(new ExcelStyleAnalyzer().Analyze(result.OutputPath)));
        var second = new ExcelStyleOptimizer().Optimize(result.OutputPath, PathOf("named-order-again.xlsx"),
            options: new(CollectUnusedStyles: gc));
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"),
            OoxmlWorkbookFixture.ReadPart(second.OutputPath, "xl/styles.xml"));
    }

    private static string WorksheetWithoutStyleReferences(string path)
    {
        var sheet = Part(path, "xl/worksheets/sheet1.xml");
        foreach (var element in sheet.Descendants())
        {
            if (element.Name == Ns + "c" || element.Name == Ns + "row") element.Attribute("s")?.Remove();
            if (element.Name == Ns + "col") element.Attribute("style")?.Remove();
        }
        return sheet.ToString(SaveOptions.DisableFormatting);
    }

    private static string[] Usage(ExcelStyleAnalysis analysis) => analysis.Worksheets.Single().Styles
        .GroupBy(s => s.Style with { StyleId = 0 })
        .Select(g => g.Key + ": " + g.Sum(s => s.CellCount))
        .OrderBy(s => s, StringComparer.Ordinal).ToArray();
}
