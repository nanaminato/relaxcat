using System.IO.Compression;
using RelaxCat.Core.Excel;
using Xunit;

namespace RelaxCat.Core.Tests;

public sealed class PrivateStyleDiagnosticsTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "secret-path-" + Guid.NewGuid().ToString("N"));
    public PrivateStyleDiagnosticsTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReportsEveryDetectedGateWithoutLeakingDocumentMetadata(bool destructive)
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>")).Replace("name=\"Normal\"", "name=\"secret-style\"")
            .Replace("</styleSheet>", "<extLst><ext uri='secret-extension-uri'><private:payload xmlns:private='secret-namespace'/></ext></extLst></styleSheet>");
        var sheet = OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1' t='str'><f>\"secret-formula\"</f><v>secret-value</v></c></row>")
            .Replace("</worksheet>", "<extLst><ext uri='secret-external-link'><private:data xmlns:private='secret-schema'/></ext></extLst></worksheet>");
        var path = OoxmlWorkbookFixture.Create(Path.Combine(directory, "secret-file.xlsx"), styles, new SheetSpec("secret-sheet", sheet));
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        using (var writer = new StreamWriter(archive.CreateEntry("xl/secret-unknown-part.xml").Open()))
            writer.Write("<private/>");
        var analysis = new ExcelStyleAnalyzer().Analyze(path, options: new ExcelStyleOptimizationOptions(DestructiveStyleCleanup: destructive));
        var report = PrivateStyleDiagnostics.FromAnalysis(analysis);
        Assert.True(report.RemappingBlocked);
        Assert.Contains("styleExtensions", report.BlockerCategories);
        Assert.Contains("unsupportedParts", report.BlockerCategories);
        Assert.Contains("packageExtensions", report.BlockerCategories);
        Assert.DoesNotContain("secret-", report.ToJson());
        Assert.Equal(analysis.StylesXmlBytes, report.StylesXmlBytes);
        Assert.Equal(report.BlockerCategories.Count, report.BlockerCategories.Distinct().Count());
    }

    [Fact]
    public void UnblockedWorkbookDoesNotInventRestrictions()
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>"));
        var path = OoxmlWorkbookFixture.Create(Path.Combine(directory, "secret-file.xlsx"), styles,
            new SheetSpec("secret-sheet", OoxmlWorkbookFixture.Worksheet("A1", "<row r='1'><c r='A1'><v>42</v></c></row>")));
        var report = PrivateStyleDiagnostics.FromAnalysis(new ExcelStyleAnalyzer().Analyze(path));
        Assert.False(report.RemappingBlocked);
        Assert.Empty(report.BlockerCategories);
        Assert.Empty(report.RetainedTableCategories);
        Assert.DoesNotContain("secret-", report.ToJson());
    }

    [Fact]
    public void FutureUnknownReasonAndRetentionDetailsCannotEscapeAllowlist()
    {
        var analysis = new ExcelStyleAnalysis("secret-path", 10, false, 1, 1, 1, 1, 1, 1, 1, 100, 90, [])
        {
            AdvancedOptimizationSkipReason = "secret-error-message",
            AdvancedOptimizationSkipReasons = ["unsupportedParts", "secret-raw-uri", "unsupportedParts"],
            RetainedStyleTables = ["secret-name", "phoneticFonts", "differentialNumberFormats"],
            Warnings = ["secret-warning"], NamedStyleCleanupCandidates = ["secret-style"]
        };
        var report = PrivateStyleDiagnostics.FromAnalysis(analysis);
        Assert.DoesNotContain("secret-", report.ToJson());
        Assert.Equal(new[] { "otherRestriction", "unsupportedParts" }, report.BlockerCategories);
        Assert.Equal(3, report.RetainedTableCategories.Count);
    }
}
