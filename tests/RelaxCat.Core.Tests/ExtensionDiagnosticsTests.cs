using System.Text.Json;
using RelaxCat.Core.Excel;
using Xunit;

namespace RelaxCat.Core.Tests;

public sealed class ExtensionDiagnosticsTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "private-path-marker-" + Guid.NewGuid().ToString("N"));
    public ExtensionDiagnosticsTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);

    private string Create(string extension, bool sheet = false)
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>",
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>"));
        if (!sheet) styles = styles.Replace("</styleSheet>", extension + "</styleSheet>");
        var worksheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' t='inlineStr'><is><t>private-cell-marker</t></is></c></row>");
        if (sheet) worksheet = worksheet.Replace("</worksheet>", extension + "</worksheet>");
        return OoxmlWorkbookFixture.Create(Path.Combine(directory, "private-filename.xlsx"), styles,
            new SheetSpec("private-sheet-marker", worksheet));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReportsOnlyTypeMetadataAndNeverUnlocksUnknownExtensions(bool sheet)
    {
        var path = Create("<extLst><ext uri='urn:synthetic'><t:payload xmlns:t='urn:test' secret='private-attribute-marker' name='private-style-marker'><t:child>private-text-marker</t:child></t:payload></ext></extLst>", sheet);
        var analysis = new ExcelStyleAnalyzer().Analyze(path, options: new ExcelStyleOptimizationOptions { CollectUnusedStyles = true, RemoveCopiedNamedStyles = true });
        var report = JsonSerializer.Serialize(analysis.ExtensionDiagnostics);
        Assert.Contains("urn:synthetic", report);
        Assert.Contains("urn:test", report);
        Assert.Contains("payload", report);
        Assert.Contains("secret", report);
        foreach (var marker in new[] { "private-attribute-marker", "private-style-marker", "private-text-marker", "private-cell-marker", "private-sheet-marker", "private-path-marker", "private-filename" })
            Assert.DoesNotContain(marker, report);
        Assert.NotNull(analysis.AdvancedOptimizationSkipReason);
        Assert.Equal(analysis.CellXfCount, analysis.ProjectedCellXfCount);
        Assert.False(analysis.ExtensionDiagnostics.Truncated);
    }

    [Fact]
    public void MetadataIsBoundedAndControlCharactersCannotForgeRows()
    {
        var types = string.Concat(Enumerable.Range(0, 100).Select(i => $"<t:element{i} attribute{i}='private-value'/ >".Replace("/ >", "/>")));
        var analysis = new ExcelStyleAnalyzer().Analyze(Create($"<extLst><ext uri='urn:synthetic&#10;&#x202e;&#x2028;&#x2029;{new string('x', 200)}'><t:payload xmlns:t='urn:test'>{types}</t:payload></ext></extLst>"));
        Assert.True(analysis.ExtensionDiagnostics.Truncated);
        Assert.InRange(analysis.ExtensionDiagnostics.Entries.Count, 1, ExtensionTypeDiagnostics.MaximumEntries);
        Assert.All(analysis.ExtensionDiagnostics.Entries, entry =>
        {
            Assert.InRange(entry.Name.Length, 1, ExtensionTypeDiagnostics.MaximumNameLength);
            Assert.DoesNotContain('\n', entry.Name);
            Assert.DoesNotContain('\u202e', entry.Name);
            Assert.DoesNotContain('\u2028', entry.Name);
            Assert.DoesNotContain('\u2029', entry.Name);
        });
    }

    [Fact]
    public void ForeignAttributesAndNamespaceDeclarationsAreVisibleWithoutTheirValues()
    {
        var path = Create("");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/styles.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            xml = xml.Replace("<styleSheet", "<styleSheet xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:custom='urn:custom' mc:Ignorable='private-prefix-marker' custom:flag='private-attribute-marker'");
            using var writer = new StreamWriter(zip.CreateEntry("xl/styles.xml").Open());
            writer.Write(xml);
        }
        var analysis = new ExcelStyleAnalyzer().Analyze(path);
        var report = JsonSerializer.Serialize(analysis.ExtensionDiagnostics);
        Assert.Contains("Ignorable", report);
        Assert.Contains("urn:custom", report);
        Assert.Contains("flag", report);
        Assert.DoesNotContain("private-prefix-marker", report);
        Assert.DoesNotContain("private-attribute-marker", report);
        Assert.NotNull(analysis.AdvancedOptimizationSkipReason);
    }

    [Fact]
    public void NoExtensionsProducesEmptyMetadata()
    {
        var report = new ExcelStyleAnalyzer().Analyze(Create("")).ExtensionDiagnostics;
        Assert.Empty(report.Entries);
        Assert.False(report.Truncated);
    }
}
