using System.Text;
using System.Xml.Linq;
using RelaxCat.Core.Excel;

namespace RelaxCat.Core.Tests;

public sealed class AncillaryStylePartTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-ancillary-" + Guid.NewGuid().ToString("N"));
    public AncillaryStylePartTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrdinaryPartsPreserveBytesAndPinReferencedNumberFormats(bool destructive)
    {
        var path = Path.Combine(directory, "in.xlsx");
        var ns = OoxmlWorkbookFixture.MainNamespace;
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf fontId='0' fillId='0' borderId='0' numFmtId='0' xfId='0'/>",
            "<xf fontId='0' fillId='0' borderId='0' numFmtId='0' xfId='0'/>"),
            numberFormats: "<numFmts count='1'><numFmt numFmtId='164' formatCode='0.0000'/></numFmts>");
        OoxmlWorkbookFixture.Create(path, styles, new SheetSpec("Synthetic", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' s='1'><v>42</v></c></row>")));
        var parts = new Dictionary<string, string>
        {
            ["xl/tables/table1.xml"] = $"<table xmlns='{ns}' id='1' name='Data' ref='A1:A2'><tableColumns count='1'><tableColumn id='1' name='Value'/></tableColumns></table>",
            ["xl/pivotTables/pivotTable1.xml"] = $"<pivotTableDefinition xmlns='{ns}'><pivotFields count='1'><pivotField numFmtId='164'/></pivotFields></pivotTableDefinition>",
            ["xl/drawings/drawing1.xml"] = "<xdr:wsDr xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><xdr:twoCellAnchor><xdr:sp><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface='Microsoft YaHei'/></a:rPr><a:t>Synthetic title</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>",
            ["xl/worksheets/_rels/sheet1.xml.rels"] = "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='h' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.invalid/' TargetMode='External'/></Relationships>"
        };
        foreach (var (name, xml) in parts) OoxmlWorkbookFixture.AddPart(path, name, Encoding.UTF8.GetBytes(xml));
        var options = new ExcelStyleOptimizationOptions(CollectUnusedStyles: true, DeduplicateComponents: true, DestructiveStyleCleanup: destructive);
        var result = new ExcelStyleOptimizer().Optimize(path, Path.Combine(directory, "out.xlsx"), options: options);
        Assert.Null(result.AdvancedOptimizationSkipReason);
        Assert.Equal(1, result.OutputCellXfCount);
        foreach (var name in parts.Keys) Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, name), OoxmlWorkbookFixture.ReadPart(result.OutputPath, name));
        var outputStyles = XDocument.Parse(Encoding.UTF8.GetString(OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml")));
        Assert.Single(outputStyles.Root!.Element(XName.Get("numFmts", ns))!.Elements());
        var second = new ExcelStyleOptimizer().Optimize(result.OutputPath, Path.Combine(directory, "again.xlsx"), options: options);
        Assert.Equal(File.ReadAllBytes(result.OutputPath), File.ReadAllBytes(second.OutputPath));
    }
    [Theory]
    [InlineData("<xdr:sp xfId='0'/>")]
    [InlineData("<xdr:sp><xdr:extLst/></xdr:sp>")]
    [InlineData("<unknown xmlns='urn:unmodeled'/>")]
    public void FamiliarDrawingPartDoesNotAuthorizeUnknownConsumers(string body)
    {
        var path = Path.Combine(directory, "unknown.xlsx");
        OoxmlWorkbookFixture.Create(path, OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf fontId='0' fillId='0' borderId='0' numFmtId='0' xfId='0'/>")),
            new SheetSpec("Synthetic", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'><v>1</v></c></row>")));
        OoxmlWorkbookFixture.AddPart(path, "xl/drawings/drawing1.xml", Encoding.UTF8.GetBytes(
            "<xdr:wsDr xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'>" + body + "</xdr:wsDr>"));
        var result = new ExcelStyleOptimizer().Optimize(path, Path.Combine(directory, "blocked.xlsx"),
            options: new ExcelStyleOptimizationOptions(DestructiveStyleCleanup: true));
        Assert.Equal("unsupportedParts", result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"));
    }

    [Theory]
    [InlineData("pageStyle")]
    [InlineData("pivotTableStyle")]
    [InlineData("vacatedStyle")]
    public void PivotNamedStyleConsumersBlockDestructiveRemoval(string attribute)
    {
        var path = Path.Combine(directory, "pivot.xlsx");
        OoxmlWorkbookFixture.Create(path, OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf fontId='0' fillId='0' borderId='0' numFmtId='0' xfId='0'/>")),
            new SheetSpec("Synthetic", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'><v>1</v></c></row>")));
        OoxmlWorkbookFixture.AddPart(path, "xl/pivotTables/pivotTable1.xml", Encoding.UTF8.GetBytes(
            $"<pivotTableDefinition xmlns='{OoxmlWorkbookFixture.MainNamespace}' {attribute}='Normal'/>"));
        var result = new ExcelStyleOptimizer().Optimize(path, Path.Combine(directory, "pivot-blocked.xlsx"),
            options: new ExcelStyleOptimizationOptions(DestructiveStyleCleanup: true));
        Assert.Equal("unsupportedParts", result.AdvancedOptimizationSkipReason);
        Assert.Empty(result.TableChanges);
    }

}
