using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using RelaxCat.Core.Excel;

namespace RelaxCat.Core.Tests;

/// <summary>Synthetic fixtures only. No issue attachment or document metadata is embedded.</summary>
public sealed class OfficeMetadataStyleTests : IDisposable
{
    private static readonly XNamespace Ns = OoxmlWorkbookFixture.MainNamespace;
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Xr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private const string Rel = OoxmlWorkbookFixture.RelationshipNamespace;
    private const string Slicer = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static readonly ExcelStyleOptimizationOptions Dangerous = new(DestructiveStyleCleanup: true);
    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-office-metadata-" + Guid.NewGuid().ToString("N"));
    public OfficeMetadataStyleTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);
    private string PathOf(string name) => Path.Combine(directory, name);

    private string Create(bool phonetic = true)
    {
        const string basis = "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'/>";
        var root = XDocument.Parse(OoxmlWorkbookFixture.StylesXml(
            OoxmlWorkbookFixture.CellXfs(
                "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>",
                "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='1'/>"),
            "<fonts count='3'><font><name val='Calibri'/></font><font><name val='Unused'/></font><font><name val='Phonetic'/></font></fonts>"));
        root.Root!.SetAttributeValue(XNamespace.Xmlns + "mc", Mc);
        root.Root.SetAttributeValue((XNamespace)Mc + "Ignorable", "xr");
        root.Root.SetAttributeValue(XNamespace.Xmlns + "xr", Xr);
        root.Root.Element(Ns + "cellStyleXfs")!.ReplaceWith(XElement.Parse($"<cellStyleXfs xmlns='{Ns}' count='2'>{basis}{basis}</cellStyleXfs>"));
        root.Root.Element(Ns + "cellStyles")!.ReplaceWith(XElement.Parse($"<cellStyles xmlns='{Ns}' count='2'><cellStyle name='Normal' xfId='0' builtinId='0'/><cellStyle xmlns:xr='{Xr}' xr:uid='{{11111111-1111-1111-1111-111111111111}}' name='Synthetic_1' xfId='1'/></cellStyles>"));
        root.Root.Add(XElement.Parse($$"""
            <extLst xmlns="{{Ns}}">
              <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:slicerStyles xmlns:x14="{{Slicer}}" defaultSlicerStyle="SlicerStyleLight1"/></ext>
              <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}"><x15:timelineStyles xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>
            </extLst>
            """));
        var worksheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' s='1' t='s'><v>0</v></c></row>")
            .Replace("</worksheet>", (phonetic ? "<phoneticPr fontId='2'/>" : "") +
                $"<pageSetup xmlns:r='{Rel}' r:id='print1'/></worksheet>");
        var input = OoxmlWorkbookFixture.Create(PathOf("input.xlsx"), root.ToString(SaveOptions.DisableFormatting), new SheetSpec("Synthetic", worksheet));
        Edit(input, "xl/workbook.xml", doc =>
        {
            doc.Root!.SetAttributeValue(XNamespace.Xmlns + "x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            doc.Root!.Add(XElement.Parse($"""
                <mc:AlternateContent xmlns:mc="{Mc}" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"><mc:Choice Requires="x15"><x15ac:absPath url="file:///synthetic/"/></mc:Choice></mc:AlternateContent>
                """));
            doc.Root.Add(XElement.Parse($"<xr:revisionPtr xmlns:xr='{Xr}' revIDLastSave='0' documentId='synthetic'/>") );
            doc.Root.Add(XElement.Parse($$"""
                <extLst xmlns="{{Ns}}"><ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}"><c:calcFeatures xmlns:c="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"><c:feature name="microsoft.com:RD"/></c:calcFeatures></ext></extLst>
                """));
        });
        Add(input, "xl/sharedStrings.xml", $"<sst xmlns='{Ns}' count='1' uniqueCount='1'><si><t>synthetic sentinel</t>{(phonetic ? "<phoneticPr fontId='2'/>" : "")}</si></sst>");
        OoxmlWorkbookFixture.AddPart(input, "xl/printerSettings/printerSettings1.bin", [1, 2, 3]);
        Add(input, "xl/worksheets/_rels/sheet1.xml.rels", Relationships($"<Relationship Id='print1' Type='{Rel}/printerSettings' Target='../printerSettings/printerSettings1.bin'/>"));
        Add(input, "xl/externalLinks/externalLink1.xml", $"<externalLink xmlns='{Ns}' xmlns:r='{Rel}'><externalBook r:id='external1'><sheetNames><sheetName val='Synthetic external sheet'/></sheetNames><sheetDataSet><sheetData sheetId='0'><row r='1'><cell r='A1' t='n'><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>");
        Add(input, "xl/externalLinks/_rels/externalLink1.xml.rels", Relationships($"<Relationship Id='external1' Type='{Rel}/externalLinkPath' Target='file:///synthetic/source.xlsx' TargetMode='External'/>"));
        return input;
    }

    [Fact]
    public void DestructiveCleanupSupportsKnownOfficeMetadataAndPinsPhoneticFonts()
    {
        var input = Create();
        var original = File.ReadAllBytes(input);
        var analysis = new ExcelStyleAnalyzer().Analyze(input, options: Dangerous);
        Assert.True(analysis.ReferenceGraphComplete);
        Assert.Contains(analysis.ReferenceGraph, e => e.Source == "xl/sharedStrings.xml" && e.Target == "fonts" && e.TargetIndex == 2);
        Assert.Equal(2, analysis.Sections.Single(s => s.Name == "fonts").ReferencedCount);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("out.xlsx"), options: Dangerous);
        Assert.True(result.Verified);
        Assert.Null(result.AdvancedOptimizationSkipReason);
        Assert.Equal(1, result.OutputCellStyleXfCount);
        Assert.Equal(1, result.OutputCellXfCount);
        Assert.Equal(analysis.ProjectedStylesXmlBytes, result.OutputStylesXmlBytes);
        Assert.DoesNotContain(result.TableChanges, c => c.Name == "fonts");
        Assert.True(XNode.DeepEquals(Part(input).Root!.Element(Ns + "fonts"), Part(result.OutputPath).Root!.Element(Ns + "fonts")));
        Assert.True(XNode.DeepEquals(Part(input).Root!.Element(Ns + "extLst"), Part(result.OutputPath).Root!.Element(Ns + "extLst")));
        foreach (var part in OoxmlWorkbookFixture.PartNames(input).Where(p => p != "xl/styles.xml" && p != "xl/worksheets/sheet1.xml"))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
        Assert.Equal(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(input, "xl/worksheets/sheet1.xml")).Replace("s='1'", "s='0'"),
            Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/worksheets/sheet1.xml")));
        Assert.Equal(original, File.ReadAllBytes(input));
        var twice = new ExcelStyleOptimizer().Optimize(result.OutputPath, PathOf("twice.xlsx"), options: Dangerous);
        Assert.Empty(twice.TableChanges);
        Assert.Equal(File.ReadAllBytes(result.OutputPath), File.ReadAllBytes(twice.OutputPath));
    }

    [Fact]
    public void DefaultSafeModeStillPreservesEveryIndexForOfficeMetadata()
    {
        var input = Create();
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("safe.xlsx"));
        Assert.NotNull(result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        foreach (var part in OoxmlWorkbookFixture.PartNames(input))
            Assert.Equal(OoxmlWorkbookFixture.ReadPart(input, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
    }

    [Fact]
    public void FontsCanStillBeCollectedWhenThereAreNoPhoneticReferences()
    {
        var input = Create(phonetic: false);
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("out.xlsx"), options: Dangerous);
        Assert.Null(result.AdvancedOptimizationSkipReason);
        Assert.Single(Part(result.OutputPath).Root!.Element(Ns + "fonts")!.Elements());
    }

    [Theory]
    [InlineData("style-child")]
    [InlineData("style-attribute")]
    [InlineData("style-uri")]
    [InlineData("calc-reference")]
    [InlineData("alternate-content")]
    [InlineData("revision-reference")]
    [InlineData("external-style-reference")]
    [InlineData("printer-relationship")]
    [InlineData("external-relationship")]
    public void RecognizedNamespaceOrPartNameNeverBypassesUnknownShape(string mutation)
    {
        var input = Create();
        switch (mutation)
        {
            case "style-child": Edit(input, "xl/styles.xml", d => d.Descendants((XNamespace)Slicer + "slicerStyles").Single().Add(new XElement((XNamespace)Slicer + "slicerStyle", new XAttribute("xfId", 1)))); break;
            case "style-attribute": Edit(input, "xl/styles.xml", d => d.Descendants((XNamespace)Slicer + "slicerStyles").Single().SetAttributeValue("fontId", 2)); break;
            case "style-uri": Edit(input, "xl/styles.xml", d => d.Descendants(Ns + "ext").First().SetAttributeValue("uri", "unknown")); break;
            case "calc-reference": Edit(input, "xl/workbook.xml", d => d.Descendants().First(e => e.Name.LocalName == "feature").SetAttributeValue("xfId", 1)); break;
            case "alternate-content": Edit(input, "xl/workbook.xml", d => d.Descendants((XNamespace)Mc + "Choice").Single().Add(new XElement((XNamespace)Slicer + "unknown"))); break;
            case "revision-reference": Edit(input, "xl/workbook.xml", d => d.Descendants((XNamespace)Xr + "revisionPtr").Single().SetAttributeValue("xfId", 1)); break;
            case "external-style-reference": Edit(input, "xl/externalLinks/externalLink1.xml", d => d.Descendants(Ns + "cell").Single().SetAttributeValue("s", 1)); break;
            case "printer-relationship": Edit(input, "xl/worksheets/_rels/sheet1.xml.rels", d => d.Root!.Elements().Single().SetAttributeValue("Type", Rel + "/unknown")); break;
            case "external-relationship": Edit(input, "xl/externalLinks/_rels/externalLink1.xml.rels", d => d.Root!.Elements().Single().SetAttributeValue("TargetMode", "Internal")); break;
        }
        var result = new ExcelStyleOptimizer().Optimize(input, PathOf("out.xlsx"), options: Dangerous);
        Assert.NotNull(result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        Assert.Equal(File.ReadAllBytes(input), File.ReadAllBytes(result.OutputPath));
    }

    [Theory]
    [InlineData("-1")]
    [InlineData("3")]
    [InlineData("invalid")]
    public void InvalidPhoneticFontReferencesAreRejected(string value)
    {
        var input = Create();
        Edit(input, "xl/sharedStrings.xml", d => d.Descendants(Ns + "phoneticPr").Single().SetAttributeValue("fontId", value));
        Assert.Throws<InvalidDataException>(() => new ExcelStyleOptimizer().Optimize(input, PathOf("out.xlsx"), options: Dangerous));
        Assert.False(File.Exists(PathOf("out.xlsx")));
    }

    private static string Relationships(string entries) => $"<Relationships xmlns='{OoxmlWorkbookFixture.PackageRelationshipNamespace}'>{entries}</Relationships>";
    private static XDocument Part(string path, string part = "xl/styles.xml") => XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(path, part)));
    private static void Add(string path, string part, string text) => OoxmlWorkbookFixture.AddPart(path, part, Encoding.UTF8.GetBytes(text));
    private static void Edit(string path, string part, Action<XDocument> edit)
    {
        var document = Part(path, part);
        edit(document);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(part)!.Delete();
        using var stream = archive.CreateEntry(part).Open();
        var bytes = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        stream.Write(bytes);
    }
}
