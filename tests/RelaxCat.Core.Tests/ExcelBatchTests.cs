using RelaxCat.Core.Excel;
using Xunit;

namespace RelaxCat.Core.Tests;

public sealed class ExcelBatchTests : IDisposable
{
    private const string DefaultXf = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""";

    private readonly string directory = Path.Combine(Path.GetTempPath(), "relaxcat-excel-batch-" + Guid.NewGuid().ToString("N"));

    public ExcelBatchTests() => Directory.CreateDirectory(directory);

    public void Dispose()
    {
        try { Directory.Delete(directory, true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private string PathOf(params string[] parts) => Path.Combine(new[] { directory }.Concat(parts).ToArray());

    private static string CreateWorkbookAt(string path)
    {
        var styles = OoxmlWorkbookFixture.StylesXml(OoxmlWorkbookFixture.CellXfs(DefaultXf, DefaultXf));
        return OoxmlWorkbookFixture.Create(path, styles, new SheetSpec("Sheet1",
            OoxmlWorkbookFixture.Worksheet("A1", """<row r="1"><c r="A1"><v>1</v></c></row>""")));
    }

    [Fact]
    public void Find_DiscoversWorkbooksRecursivelyAndSkipsOutputsAndTemporaryFiles()
    {
        var input = PathOf("input");
        Directory.CreateDirectory(input);
        File.WriteAllText(Path.Combine(input, "a.xlsx"), "x");
        File.WriteAllText(Path.Combine(input, "b.XLSM"), "x");
        File.WriteAllText(Path.Combine(input, "notes.txt"), "x");
        File.WriteAllText(Path.Combine(input, "~$temp.xlsx"), "x");
        File.WriteAllText(Path.Combine(input, "already-optimized.xlsx"), "x");
        Directory.CreateDirectory(Path.Combine(input, "sub"));
        File.WriteAllText(Path.Combine(input, "sub", "c.xlsx"), "x");
        var output = Path.Combine(input, "out");
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, "inside.xlsx"), "x");

        var recursive = ExcelWorkbookFinder.Find(input, recursive: true, output);

        Assert.Equal(new[] { "a.xlsx", "b.XLSM", Path.Combine("sub", "c.xlsx") },
            recursive.Files.Select(file => file.RelativePath));
        Assert.Empty(recursive.Issues);

        var shallow = ExcelWorkbookFinder.Find(input, recursive: false, output);
        Assert.Equal(new[] { "a.xlsx", "b.XLSM" }, shallow.Files.Select(file => file.RelativePath));
        Assert.Throws<DirectoryNotFoundException>(() => ExcelWorkbookFinder.Find(PathOf("missing"), true));
    }

    [Fact]
    public void Optimize_ContinuesAfterFailuresAndSkipsExistingOutputs()
    {
        var input = PathOf("input");
        var output = PathOf("out");
        var one = CreateWorkbookAt(Path.Combine(input, "one.xlsx"));
        var two = CreateWorkbookAt(Path.Combine(input, "sub", "two.xlsx"));
        File.WriteAllText(Path.Combine(input, "bad.xlsx"), "not a zip package");
        Directory.CreateDirectory(output);
        var existing = Path.Combine(output, "one.xlsx");
        File.WriteAllText(existing, "existing");
        var oneBefore = File.ReadAllBytes(one);
        var twoBefore = File.ReadAllBytes(two);
        var discovery = ExcelWorkbookFinder.Find(input, true, output);
        var progress = new ProgressCollector<ExcelBatchProgress>();

        var outcome = new ExcelBatchOptimizer().Optimize(discovery.Files, output, overwrite: false, progress);

        Assert.False(outcome.Cancelled);
        Assert.Equal(3, outcome.Total);
        Assert.Equal(3, outcome.Results.Count);
        var byName = outcome.Results.ToDictionary(result => result.RelativePath, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(ExcelBatchItemStatus.Skipped, byName["one.xlsx"].Status);
        Assert.Equal(ExcelBatchItemStatus.Failed, byName["bad.xlsx"].Status);
        Assert.Equal(ExcelBatchItemStatus.Completed, byName[Path.Combine("sub", "two.xlsx")].Status);
        Assert.Equal("existing", File.ReadAllText(existing));
        Assert.Equal(oneBefore, File.ReadAllBytes(one));
        Assert.Equal(twoBefore, File.ReadAllBytes(two));
        var twoOutput = Path.Combine(output, "sub", "two.xlsx");
        Assert.True(File.Exists(twoOutput));
        Assert.Equal(1, new ExcelStyleAnalyzer().Analyze(twoOutput).CellXfCount);
        Assert.Equal(3, progress.Reports.Count(report => report.LastResult is not null));
        Assert.Equal(3, progress.Reports.Max(report => report.Completed));

        var oneFile = discovery.Files.Single(file => file.RelativePath == "one.xlsx");
        var overwritten = new ExcelBatchOptimizer().Optimize(new[] { oneFile }, output, overwrite: true);
        Assert.True(overwritten.Results.Single().Succeeded);
        Assert.Equal(1, new ExcelStyleAnalyzer().Analyze(existing).CellXfCount);
    }

    [Fact]
    public void Optimize_StopsOnCancellationAndKeepsCompletedResults()
    {
        var input = PathOf("input-cancel");
        CreateWorkbookAt(Path.Combine(input, "first.xlsx"));
        CreateWorkbookAt(Path.Combine(input, "second.xlsx"));
        var files = ExcelWorkbookFinder.Find(input, true).Files;
        var output = PathOf("out-cancel");
        using var cancellation = new CancellationTokenSource();
        var progress = new ProgressCollector<ExcelBatchProgress>(report =>
        {
            if (report.LastResult is not null) cancellation.Cancel();
        });

        var outcome = new ExcelBatchOptimizer().Optimize(files, output, overwrite: false, progress, cancellation.Token);

        Assert.True(outcome.Cancelled);
        Assert.Single(outcome.Results);
        Assert.Equal(ExcelBatchItemStatus.Completed, outcome.Results[0].Status);
        Assert.Equal("first.xlsx", outcome.Results[0].RelativePath);

        using var alreadyCancelled = new CancellationTokenSource();
        alreadyCancelled.Cancel();
        var empty = new ExcelBatchOptimizer().Optimize(files, PathOf("out-cancel-2"), false, null, alreadyCancelled.Token);
        Assert.True(empty.Cancelled);
        Assert.Empty(empty.Results);
    }

    [Theory]
    [InlineData("book.xlsx", false, "book.xlsx")]
    [InlineData("book.xlsm", true, "book-optimized.xlsm")]
    [InlineData("book-optimized.xlsx", true, "book-optimized.xlsx")]
    [InlineData("book-OPTIMIZED.XLSM", true, "book-OPTIMIZED.XLSM")]
    public void OutputFileName_IsOptionalAndDoesNotAccumulateSuffixes(string input, bool suffix, string expected)
        => Assert.Equal(expected, ExcelBatchPaths.OutputFileName(input, suffix));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Optimize_PreservesNamesOrAddsSuffixAndNeverChangesSource(bool suffix)
    {
        var input = CreateWorkbookAt(PathOf("input", "sub", "one.xlsx"));
        var bytes = File.ReadAllBytes(input);
        var files = ExcelWorkbookFinder.Find(PathOf("input"), true).Files;
        var output = PathOf("out");
        var result = new ExcelBatchOptimizer().Optimize(files, output, false, appendSuffix: suffix);
        Assert.True(Assert.Single(result.Results).Succeeded);
        Assert.Equal(Path.Combine(output, "sub", suffix ? "one-optimized.xlsx" : "one.xlsx"), result.Results[0].OutputPath);
        Assert.Equal(bytes, File.ReadAllBytes(input));
        Assert.Equal(ExcelBatchItemStatus.Skipped, Assert.Single(new ExcelBatchOptimizer()
            .Optimize(files, output, false, appendSuffix: suffix).Results).Status);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Optimize_RejectsSourceOverwriteEvenWhenOverwriteIsEnabled(bool overwrite)
    {
        var input = CreateWorkbookAt(PathOf("input", "one.xlsx"));
        var bytes = File.ReadAllBytes(input);
        var files = ExcelWorkbookFinder.Find(PathOf("input"), true).Files;
        Assert.Throws<ArgumentException>(() => new ExcelBatchOptimizer().Optimize(files, PathOf("input"), overwrite));
        Assert.Equal(bytes, File.ReadAllBytes(input));
    }

    [Fact]
    public void Optimize_RejectsCrossInputAndDuplicateOutputCollisionsBeforeWriting()
    {
        var one = CreateWorkbookAt(PathOf("input", "one.xlsx"));
        var two = CreateWorkbookAt(PathOf("input", "one-optimized.xlsx"));
        var files = new[] { new ExcelBatchFile(one, "one.xlsx", 0), new ExcelBatchFile(two, "one-optimized.xlsx", 0) };
        Assert.Throws<ArgumentException>(() => new ExcelBatchOptimizer().Optimize(files, PathOf("input"), true, appendSuffix: true));
        Assert.Throws<ArgumentException>(() => new ExcelBatchOptimizer().Optimize(files, PathOf("output"), true, appendSuffix: true));
        Assert.False(Directory.Exists(PathOf("output")));
    }

    private sealed class ProgressCollector<T> : IProgress<T>
    {
        private readonly Action<T>? callback;

        public ProgressCollector(Action<T>? callback = null) => this.callback = callback;

        public List<T> Reports { get; } = new();

        public void Report(T value)
        {
            Reports.Add(value);
            callback?.Invoke(value);
        }
    }
}
