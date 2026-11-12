using System.Text;
using RelaxCat.Core.Excel;

namespace RelaxCat.Core.Tests;

public sealed class UnsupportedPartDiagnosticTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
    public UnsupportedPartDiagnosticTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);
    private const string DrawingStart = "<xdr:wsDr xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'>";
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static TheoryData<string, string, UnsupportedPartKind, UnsupportedPartReason> Cases => new()
    {
        {"xl/secret-name.xml", "<secret/>", UnsupportedPartKind.Other, UnsupportedPartReason.UnrecognizedPart},
        {"xl/drawings/drawing1.xml", "<secret", UnsupportedPartKind.Drawing, UnsupportedPartReason.MalformedXml},
        {"xl/drawings/drawing1.xml", "<secret/>", UnsupportedPartKind.Drawing, UnsupportedPartReason.UnexpectedRoot},
        {"xl/drawings/drawing1.xml", DrawingStart + "<xdr:extLst/></xdr:wsDr>", UnsupportedPartKind.Drawing, UnsupportedPartReason.ExtensionList},
        {"xl/drawings/drawing1.xml", DrawingStart + "<AlternateContent xmlns='http://schemas.openxmlformats.org/markup-compatibility/2006'/></xdr:wsDr>", UnsupportedPartKind.Drawing, UnsupportedPartReason.AlternateContent},
        {"xl/drawings/drawing1.xml", DrawingStart + "<secret xmlns='secret-namespace'/></xdr:wsDr>", UnsupportedPartKind.Drawing, UnsupportedPartReason.UnknownElementNamespace},
        {"xl/drawings/drawing1.xml", DrawingStart + "<xdr:sp xmlns:s='secret-namespace' s:secret='secret-value'/></xdr:wsDr>", UnsupportedPartKind.Drawing, UnsupportedPartReason.UnknownAttributeNamespace},
        {"xl/drawings/drawing1.xml", DrawingStart + "<xdr:sp xfId='0'/></xdr:wsDr>", UnsupportedPartKind.Drawing, UnsupportedPartReason.UnhandledStyleReference},
        {"xl/tables/table1.xml", "<table xmlns='" + Main + "' dxfId='999'/>", UnsupportedPartKind.Table, UnsupportedPartReason.InvalidStyleReference},
        {"xl/tables/table1.xml", "<table xmlns='" + Main + "' dataCellStyle='secret-name'/>", UnsupportedPartKind.Table, UnsupportedPartReason.NamedStyleReference},
        {"xl/drawings/secret.vml", "<xml><ClientData xmlns='urn:schemas-microsoft-com:office:excel' ObjectType='Text'/></xml>", UnsupportedPartKind.VmlDrawing, UnsupportedPartReason.VmlNonCommentObject},
        {"xl/drawings/secret.vml", "<xml><OLEObject xmlns='urn:schemas-microsoft-com:office:office'/></xml>", UnsupportedPartKind.VmlDrawing, UnsupportedPartReason.VmlUnsupportedObject},
        {"xl/secret/_rels/secret.xml.rels", "<secret/>", UnsupportedPartKind.Relationship, UnsupportedPartReason.RelationshipOwner},
        {"xl/worksheets/_rels/sheet1.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='secret' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/secret' Target='secret'/></Relationships>", UnsupportedPartKind.Relationship, UnsupportedPartReason.RelationshipType},
        {"xl/worksheets/_rels/sheet1.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='secret' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='secret-missing'/></Relationships>", UnsupportedPartKind.Relationship, UnsupportedPartReason.InvalidRelationship}
    };

    private string Create()
    {
        var path = Path.Combine(directory, Guid.NewGuid() + ".xlsx");
        return OoxmlWorkbookFixture.Create(path, OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf fontId='0' fillId='0' borderId='0' numFmtId='0' xfId='0'/>")),
            new SheetSpec("secret-sheet", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' t='str'><v>secret-content</v></c></row>")));
    }

    [Theory, MemberData(nameof(Cases))]
    public void FirstRejectionHasFixedCategoryAndPreservesFailClosed(string part, string xml, UnsupportedPartKind kind, UnsupportedPartReason reason)
    {
        var path = Create();
        OoxmlWorkbookFixture.AddPart(path, part, Encoding.UTF8.GetBytes(xml));
        var options = new ExcelStyleOptimizationOptions(DestructiveStyleCleanup: true);
        var analysis = new ExcelStyleAnalyzer().Analyze(path, options: options);
        var report = PrivateStyleDiagnostics.FromAnalysis(analysis);
        Assert.True(report.RemappingBlocked);
        Assert.Equal(new UnsupportedPartDiagnostic(kind, reason, 1), Assert.Single(report.UnsupportedPartDetails));
        Assert.DoesNotContain("secret", report.ToJson(), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(part, report.ToJson());
        var result = new ExcelStyleOptimizer().Optimize(path, Path.Combine(directory, Guid.NewGuid() + ".xlsx"), options: options);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, "xl/styles.xml"), OoxmlWorkbookFixture.ReadPart(result.OutputPath, "xl/styles.xml"));
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, part), OoxmlWorkbookFixture.ReadPart(result.OutputPath, part));
    }

    [Theory]
    [InlineData("valid", false)]
    [InlineData("extraAttribute", true)]
    [InlineData("extraElement", true)]
    [InlineData("wrongUri", true)]
    [InlineData("wrongNamespace", true)]
    [InlineData("wrongParent", true)]
    [InlineData("invalidGuid", true)]
    [InlineData("text", true)]
    public void CreationIdentityIsNarrowlySupportedAndPreserved(string variation, bool blocked)
    {
        var path = Create();
        var xml = DrawingStart + "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id='1' name='secret-title'>" +
            "<a:extLst xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:ext uri='{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'>" +
            "<a16:creationId xmlns:a16='http://schemas.microsoft.com/office/drawing/2014/main' id='{24C4C274-5652-432E-8F41-76932F11AACE}'/>" +
            "</a:ext></a:extLst></xdr:cNvPr></xdr:nvSpPr></xdr:sp></xdr:wsDr>";
        xml = variation switch
        {
            "extraAttribute" => xml.Replace("<a16:creationId ", "<a16:creationId fontId='0' "),
            "extraElement" => xml.Replace("</a:ext>", "<a:secret/></a:ext>"),
            "wrongUri" => xml.Replace("FF2B5EF4-FFF2-40B4-BE49-F238E27FC236", "00000000-0000-0000-0000-000000000000"),
            "wrongNamespace" => xml.Replace("http://schemas.microsoft.com/office/drawing/2014/main", "secret-namespace"),
            "wrongParent" => xml.Replace("xdr:cNvPr", "xdr:spPr"),
            "invalidGuid" => xml.Replace("24C4C274-5652-432E-8F41-76932F11AACE", "secret-invalid"),
            "text" => xml.Replace("</a:ext>", "secret-content</a:ext>"),
            _ => xml
        };
        OoxmlWorkbookFixture.AddPart(path, "xl/drawings/drawing1.xml", Encoding.UTF8.GetBytes(xml));
        var options = new ExcelStyleOptimizationOptions(DestructiveStyleCleanup: true);
        var report = PrivateStyleDiagnostics.FromAnalysis(new ExcelStyleAnalyzer().Analyze(path, options: options));
        Assert.Equal(blocked, report.RemappingBlocked);
        Assert.DoesNotContain("secret", report.ToJson(), StringComparison.OrdinalIgnoreCase);
        var output = new ExcelStyleOptimizer().Optimize(path, Path.Combine(directory, Guid.NewGuid() + ".xlsx"), options: options);
        Assert.Equal(OoxmlWorkbookFixture.ReadPart(path, "xl/drawings/drawing1.xml"), OoxmlWorkbookFixture.ReadPart(output.OutputPath, "xl/drawings/drawing1.xml"));
        var second = new ExcelStyleOptimizer().Optimize(output.OutputPath, Path.Combine(directory, Guid.NewGuid() + ".xlsx"), options: options);
        Assert.Equal(File.ReadAllBytes(output.OutputPath), File.ReadAllBytes(second.OutputPath));
    }

    [Fact]
    public void RepeatedPartsAreAggregatedWithoutNamesOrUnboundedList()
    {
        var path = Create();
        for (var i = 1; i <= 100; i++)
            OoxmlWorkbookFixture.AddPart(path, $"xl/drawings/drawing{i}.xml", Encoding.UTF8.GetBytes(DrawingStart + "<xdr:extLst/></xdr:wsDr>"));
        var report = PrivateStyleDiagnostics.FromAnalysis(new ExcelStyleAnalyzer().Analyze(path));
        Assert.Equal(new UnsupportedPartDiagnostic(UnsupportedPartKind.Drawing, UnsupportedPartReason.ExtensionList, 100), Assert.Single(report.UnsupportedPartDetails));
    }
}
