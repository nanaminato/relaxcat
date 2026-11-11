using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using RelaxCat.Core.Excel;

namespace RelaxCat.Core.Tests;

/// <summary>Every workbook is generated here. Lossy output is never asserted semantically equivalent.</summary>
public sealed class DestructiveStyleTests : IDisposable
{
    private static readonly XNamespace Ns = OoxmlWorkbookFixture.MainNamespace;
    private static readonly ExcelStyleOptimizationOptions Dangerous = new(DestructiveStyleCleanup: true);
    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-destructive-" + Guid.NewGuid().ToString("N"));
    public DestructiveStyleTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);
    private string PathOf(string name) => Path.Combine(directory, name);
    private static string Base(int font = 0) => $"<xf numFmtId='0' fontId='{font}' fillId='0' borderId='0'/>";
    private static string Cell(int basis = 0, int font = 0, string extra = "") =>
        $"<xf numFmtId='0' fontId='{font}' fillId='0' borderId='0' xfId='{basis}' {extra}/>";
    private static string Name(string name, int basis, string extra = "") => $"<cellStyle name='{name}' xfId='{basis}' {extra}/>";
    private string Create(string[] bases, string[] cells, string[] names, string? worksheet = null)
    {
        var document = XDocument.Parse(OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(cells),
            "<fonts count='3'><font><name val='Calibri'/></font><font><name val='Arial'/><b/></font><font><name val='Unused'/></font></fonts>"));
        document.Root!.Element(Ns + "cellStyleXfs")!.ReplaceWith(XElement.Parse($"<cellStyleXfs xmlns='{Ns}' count='{bases.Length}'>{string.Concat(bases)}</cellStyleXfs>"));
        document.Root.Element(Ns + "cellStyles")!.ReplaceWith(XElement.Parse($"<cellStyles xmlns='{Ns}' count='{names.Length}'>{string.Concat(names)}</cellStyles>"));
        return OoxmlWorkbookFixture.Create(PathOf("input.xlsx"), document.ToString(SaveOptions.DisableFormatting),
            new SheetSpec("Synthetic", worksheet ?? OoxmlWorkbookFixture.Worksheet("A1:C1",
                "<row r='1'><c r='A1' s='0'><f>SUM(B1:C1)</f><v>5</v></c><c r='B1'><v>2</v></c><c r='C1'><v>3</v></c></row>")));
    }
    private static XDocument Part(string path, string part = "xl/styles.xml") =>
        XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(path, part)));
    private static XElement[] Table(string path, string name) => Part(path).Root!.Element(Ns + name)!.Elements().ToArray();
    private ExcelOptimizationResult Optimize(string path, string output = "out.xlsx") => new ExcelStyleOptimizer().Optimize(path, PathOf(output), options: Dangerous);

    [Fact]
    public void FortyEightThousandNamedIdentitiesCollapseInsteadOfBeingCalledDuplicates()
    {
        const int count = 48846;
        var path = Create(Enumerable.Range(0, count).Select(_ => Base()).ToArray(),
            [Cell(), Cell(count - 1)], Enumerable.Range(0, count).Select(i => Name(i == 0 ? "Normal" : "Imported_" + i, i, i == 0 ? "builtinId='0'" : "hidden='1'")).ToArray(),
            OoxmlWorkbookFixture.Worksheet("A1:B1", "<row r='1'><c r='A1' s='1'><v>7</v></c><c r='B1'/></row>"));
        var original = File.ReadAllBytes(path);
        var analysis = new ExcelStyleAnalyzer().Analyze(path, options: Dangerous);
        var result = Optimize(path);
        Assert.Equal(1, result.OutputCellStyleXfCount);
        Assert.Equal(1, result.OutputCellXfCount);
        Assert.Single(Table(result.OutputPath, "cellStyles"));
        Assert.True(result.OutputStylesXmlBytes < result.InputStylesXmlBytes / 1000);
        Assert.Equal(result.OutputStylesXmlBytes, analysis.ProjectedStylesXmlBytes);
        Assert.Equal(count - 1, result.TableChanges.Single(c => c.Name == "cellStyles").DestructiveRemoved);
        Assert.Equal(0, result.TableChanges.Single(c => c.Name == "cellStyles").DuplicatesRemoved);
        Assert.Equal(count, result.NamedStyleMap.Count);
        Assert.Equal(0, result.NamedStyleMap[0]);
        Assert.All(result.NamedStyleMap.Skip(1), id => Assert.Equal(-1, id));
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.True(result.IsLossy);
        Assert.True(result.Verified);
        Assert.Contains("NotVisualOrSemanticEquivalence", result.VerificationScope);
        Assert.NotEmpty(result.Warnings);
        Assert.All(analysis.Sections, section => Assert.Equal(0, section.SafeDuplicateRemovalCount));
    }

    [Fact]
    public void PreservesLocalizedNormalAtNonzeroBaseAndItsFontMetrics()
    {
        var path = Create([Base(), Base(1)], [Cell(1)], [Name("Discard", 0), Name("標準", 1, "builtinId='0'")]);
        var result = Optimize(path);
        Assert.Equal("標準", Table(result.OutputPath, "cellStyles")[0].Attribute("name")!.Value);
        var normal = Table(result.OutputPath, "cellStyleXfs")[0];
        var font = Table(result.OutputPath, "fonts")[int.Parse(normal.Attribute("fontId")!.Value)];
        Assert.Equal("Arial", font.Element(Ns + "name")!.Attribute("val")!.Value);
        Assert.Equal(new[] { -1, 0 }, result.CellStyleXfMap);
        Assert.Equal(new[] { -1, 0 }, result.NamedStyleMap);
    }

    [Fact]
    public void ApplyFalseDoesNotReplaceDirectFontAlignmentProtectionOrNumberFormatWithBase()
    {
        var direct = "<xf numFmtId='14' fontId='1' fillId='0' borderId='0' xfId='0' applyFont='false' applyAlignment='0' applyProtection='0' applyNumberFormat='0' quotePrefix='true'><alignment horizontal='right' wrapText='true'/><protection locked='false' hidden='true'/></xf>";
        var path = Create([Base()], [direct], [Name("Normal", 0, "builtinId='0'")]);
        var result = Optimize(path);
        var xf = Table(result.OutputPath, "cellXfs")[0];
        Assert.Equal("1", xf.Attribute("fontId")!.Value);
        Assert.Equal("14", xf.Attribute("numFmtId")!.Value);
        Assert.Equal("right", xf.Element(Ns + "alignment")!.Attribute("horizontal")!.Value);
        Assert.Equal("0", xf.Element(Ns + "protection")!.Attribute("locked")!.Value);
        Assert.Equal("1", xf.Element(Ns + "protection")!.Attribute("hidden")!.Value);
        Assert.Equal("1", xf.Attribute("quotePrefix")!.Value);
        Assert.All(new[] { "applyFont", "applyAlignment", "applyProtection", "applyNumberFormat" }, flag => Assert.Equal("1", xf.Attribute(flag)!.Value));
    }

    [Fact]
    public void MissingDirectComponentsAreNotInventedFromBase()
    {
        var path = Create([Base(1)], ["<xf xfId='0'/>", Cell()], [Name("Normal", 0, "builtinId='0'")]);
        var result = Optimize(path);
        Assert.Null(Table(result.OutputPath, "cellXfs")[0].Attribute("fontId"));
        Assert.Null(Table(result.OutputPath, "cellXfs")[0].Element(Ns + "alignment"));
    }

    [Fact]
    public void RowColumnCellMappingsPreserveFormulaCachedValuesMergesAndEveryOtherByte()
    {
        var sheet = $"<worksheet xmlns='{Ns}'><dimension ref='A1:D2'/><cols><col min='4' max='4' style='3'/></cols><sheetData><row r='1' s='2' customFormat='1'><c r='A1' s='1'><f>SUM(B1:C1)</f><v>5</v></c><c r='B1'><v>2</v></c><c r='C1'><v>3</v></c><c r='D1'/></row><row r='2'><c r='A2' t='inlineStr'><is><t>sentinel</t></is></c></row></sheetData><mergeCells count='1'><mergeCell ref='A2:C2'/></mergeCells></worksheet>";
        var path = Create([Base(), Base()], [Cell(), Cell(1), Cell(1), Cell()], [Name("Normal", 0, "builtinId='0'"), Name("Custom", 1)], sheet);
        var result = Optimize(path);
        Assert.Equal(new[] { 0, 0, 0, 0 }, result.CellXfMap);
        var expected = sheet.Replace("style='3'", "style='0'").Replace("s='2'", "s='0'").Replace("s='1'", "s='0'");
        Assert.Equal(expected, Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/worksheets/sheet1.xml")));
        foreach (var part in OoxmlWorkbookFixture.PartNames(path).Where(p => p is not "xl/styles.xml" and not "xl/worksheets/sheet1.xml"))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
    }

    [Fact]
    public void SecondPassIsByteIdenticalAndHasNoFurtherChanges()
    {
        var path = Create([Base(), Base(1)], [Cell(), Cell(1)], [Name("Normal", 0, "builtinId='0'"), Name("Gone", 1)]);
        var first = Optimize(path);
        var second = Optimize(first.OutputPath, "twice.xlsx");
        Assert.Empty(second.TableChanges);
        Assert.Equal(File.ReadAllBytes(first.OutputPath), File.ReadAllBytes(second.OutputPath));
    }

    [Fact]
    public void DuplicateNamesAndBaseOwnersCanBeRemovedOnlyByExplicitDestructiveMode()
    {
        var path = Create([Base()], [Cell()], [Name("Normal", 0, "builtinId='0'"), Name("Normal", 0)]);
        Assert.Throws<InvalidDataException>(() => new ExcelStyleOptimizer().Optimize(path, PathOf("safe.xlsx")));
        var result = Optimize(path);
        Assert.Single(Table(result.OutputPath, "cellStyles"));
        Assert.True(new ExcelStyleAnalyzer().Analyze(result.OutputPath).ReferenceGraphComplete);
    }

    [Fact]
    public void AbsentNormalIsRecreatedWithoutRetainingCustomNames()
    {
        var path = Create([Base()], [Cell()], [Name("Custom", 0)]);
        var result = Optimize(path);
        Assert.Equal("Normal", Table(result.OutputPath, "cellStyles")[0].Attribute("name")!.Value);
        Assert.Equal(new[] { -1 }, result.NamedStyleMap);
    }

    [Fact]
    public void VbaIsOnlyOptedInAsBehaviorRiskAndItsBytesArePreserved()
    {
        var path = Create([Base(), Base()], [Cell()], [Name("Normal", 0, "builtinId='0'"), Name("UsedByMacro", 1)]);
        var payload = Encoding.UTF8.GetBytes("synthetic binary sentinel, not real executable VBA");
        OoxmlWorkbookFixture.AddPart(path, "xl/vbaProject.bin", payload);
        var safe = new ExcelStyleOptimizer().Optimize(path, PathOf("safe.xlsx"));
        Assert.NotNull(safe.AdvancedOptimizationSkipReason);
        var result = Optimize(path);
        Assert.Null(result.AdvancedOptimizationSkipReason);
        Assert.Single(Table(result.OutputPath, "cellStyles"));
        Assert.Equal(payload, OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/vbaProject.bin"));
    }

    [Theory]
    [InlineData("xl/unknown.xml")]
    [InlineData("_xmlsignatures/sig1.xml")]
    public void UnknownPartsAndSignaturesStillBlockIndexRewriting(string part)
    {
        var path = Create([Base(), Base()], [Cell()], [Name("Normal", 0, "builtinId='0'"), Name("Keep", 1)]);
        OoxmlWorkbookFixture.AddPart(path, part, Encoding.UTF8.GetBytes("<unknown xfId='1'/>"));
        var result = Optimize(path);
        Assert.NotNull(result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"));
    }

    [Theory]
    [InlineData("<xf fontId='999' xfId='0'/>")]
    [InlineData("<xf fontId='0' xfId='999'/>")]
    public void DestructiveModeStillRejectsInvalidKnownReferences(string cell)
    {
        var path = Create([Base()], [cell], [Name("Normal", 0, "builtinId='0'")]);
        Assert.Throws<InvalidDataException>(() => Optimize(path));
        Assert.False(File.Exists(PathOf("out.xlsx")));
    }

    [Fact]
    public void SourceCannotBeOverwrittenEvenWithDestructiveMode()
    {
        var path = Create([Base()], [Cell()], [Name("Normal", 0, "builtinId='0'")]);
        Assert.Throws<ArgumentException>(() => new ExcelStyleOptimizer().Optimize(path, path, options: Dangerous));
    }

    [Fact]
    public void OptionalComponentAndBaseTablesCanAllBeAbsent()
    {
        var path = OoxmlWorkbookFixture.Create(PathOf("input.xlsx"),
            $"<styleSheet xmlns='{Ns}'><cellXfs count='1'><xf/></cellXfs></styleSheet>",
            new SheetSpec("Minimal", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'><v>7</v></c></row>")));
        var result = Optimize(path);
        Assert.True(result.Verified);
        Assert.Single(Table(result.OutputPath, "cellStyleXfs"));
        Assert.False(Table(result.OutputPath, "cellStyleXfs")[0].HasAttributes);
        Assert.True(new ExcelStyleAnalyzer().Analyze(result.OutputPath).ReferenceGraphComplete);
    }

    [Fact]
    public void UnknownStyleExtensionsAreNotCopiedWithStaleIndices()
    {
        var path = Create([Base(), Base()], [Cell()], [Name("Normal", 0, "builtinId='0'"), Name("Keep", 1)]);
        var document = Part(path);
        document.Root!.Add(new XElement(Ns + "extLst", new XElement(Ns + "ext", new XAttribute("uri", "synthetic"),
            new XElement(XName.Get("unknown", "urn:test"), new XAttribute("xfId", 1)))));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            zip.GetEntry("xl/styles.xml")!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("xl/styles.xml").Open(), new UTF8Encoding(false));
            writer.Write(document.ToString(SaveOptions.DisableFormatting));
        }
        var result = Optimize(path);
        Assert.NotNull(result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"));
    }
}
