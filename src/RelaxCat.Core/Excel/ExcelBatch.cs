namespace RelaxCat.Core.Excel;

public sealed record ExcelBatchFile(string Path, string RelativePath, long SizeBytes);

public sealed record ExcelBatchDiscovery(IReadOnlyList<ExcelBatchFile> Files, IReadOnlyList<string> Issues);

public enum ExcelBatchItemStatus { Completed, Skipped, Failed, Cancelled }

public sealed record ExcelBatchItemResult(
    string InputPath,
    string RelativePath,
    string? OutputPath,
    ExcelBatchItemStatus Status,
    int InputCellXfCount,
    int OutputCellXfCount,
    long InputBytes,
    long OutputBytes,
    bool Verified,
    string? Error)
{
    public IReadOnlyList<StyleTableChange> TableChanges { get; init; } = [];
    public string? AdvancedOptimizationSkipReason { get; init; }
    public bool Succeeded => Status == ExcelBatchItemStatus.Completed;
}

public sealed record ExcelBatchProgress(
    int Completed,
    int Total,
    string? CurrentFile,
    double FileProgress,
    ExcelBatchItemResult? LastResult);

public sealed record ExcelBatchOutcome(IReadOnlyList<ExcelBatchItemResult> Results, bool Cancelled, int Total);

public static class ExcelBatchPaths
{
    public const string OutputSuffix = "-optimized";

    public static string OutputPathFor(string outputRoot, ExcelBatchFile file)
    {
        var directory = Path.Combine(outputRoot, Path.GetDirectoryName(file.RelativePath) ?? "");
        var name = Path.GetFileNameWithoutExtension(file.Path) + OutputSuffix + Path.GetExtension(file.Path);
        return Path.Combine(directory, name);
    }
}

/// <summary>Recursively finds .xlsx/.xlsm workbooks without following reparse points.</summary>
public static class ExcelWorkbookFinder
{
    public static ExcelBatchDiscovery Find(string folder, bool recursive, string? outputFolder = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(folder))
            throw new ArgumentException("An input folder is required.", nameof(folder));
        if (!Directory.Exists(folder))
            throw new DirectoryNotFoundException("The input folder does not exist: " + folder);
        var root = Path.GetFullPath(folder);
        var output = string.IsNullOrWhiteSpace(outputFolder) ? null : Path.GetFullPath(outputFolder);
        var files = new List<ExcelBatchFile>();
        var issues = new List<string>();
        var pending = new Stack<string>();
        pending.Push(root);
        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();
            string[] entries;
            try { entries = Directory.GetFileSystemEntries(current); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException)
            {
                issues.Add(current);
                continue;
            }
            foreach (var entry in entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                FileAttributes attributes;
                try { attributes = File.GetAttributes(entry); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                {
                    issues.Add(entry);
                    continue;
                }
                if (attributes.HasFlag(FileAttributes.ReparsePoint)) continue;
                if (attributes.HasFlag(FileAttributes.Directory))
                {
                    // A generated output folder inside the input tree must not be scanned again.
                    if (output is not null && !IsSame(output, root) && IsSameOrInside(output, entry)) continue;
                    if (recursive) pending.Push(entry);
                    continue;
                }
                if (!ExcelStyleAnalyzer.IsSupported(entry)) continue;
                var name = Path.GetFileName(entry);
                if (name.StartsWith("~$", StringComparison.Ordinal)) continue;
                if (name.EndsWith(ExcelBatchPaths.OutputSuffix + Path.GetExtension(entry), StringComparison.OrdinalIgnoreCase))
                    continue;
                long size;
                try { size = new FileInfo(entry).Length; }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                {
                    issues.Add(entry);
                    continue;
                }
                files.Add(new ExcelBatchFile(entry, Path.GetRelativePath(root, entry), size));
            }
        }
        files.Sort((left, right) => string.Compare(left.RelativePath, right.RelativePath, StringComparison.OrdinalIgnoreCase));
        return new ExcelBatchDiscovery(files, issues);
    }

    private static bool IsSame(string left, string right) =>
        string.Equals(Path.TrimEndingDirectorySeparator(Path.GetFullPath(left)),
            Path.TrimEndingDirectorySeparator(Path.GetFullPath(right)), StringComparison.OrdinalIgnoreCase);

    private static bool IsSameOrInside(string parent, string child)
    {
        var normalizedParent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(parent)) + Path.DirectorySeparatorChar;
        var normalizedChild = Path.GetFullPath(child);
        return normalizedChild.StartsWith(normalizedParent, StringComparison.OrdinalIgnoreCase) ||
               IsSame(parent, child);
    }
}

/// <summary>Runs the single-workbook optimizer over a discovered file list, continuing after per-file failures.</summary>
public sealed class ExcelBatchOptimizer
{
    public ExcelBatchOutcome Optimize(IReadOnlyList<ExcelBatchFile> files, string outputRoot, bool overwrite,
        IProgress<ExcelBatchProgress>? progress = null, CancellationToken cancellationToken = default,
        ExcelStyleOptimizationOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(outputRoot))
            throw new ArgumentException("An output folder is required.", nameof(outputRoot));
        var root = Path.GetFullPath(outputRoot);
        var results = new List<ExcelBatchItemResult>(files.Count);
        if (cancellationToken.IsCancellationRequested) return new ExcelBatchOutcome(results, true, files.Count);
        for (var index = 0; index < files.Count; index++)
        {
            if (cancellationToken.IsCancellationRequested) return new ExcelBatchOutcome(results, true, files.Count);
            var file = files[index];
            var output = ExcelBatchPaths.OutputPathFor(root, file);
            ExcelBatchItemResult result;
            if (!overwrite && File.Exists(output))
            {
                result = new ExcelBatchItemResult(file.Path, file.RelativePath, output, ExcelBatchItemStatus.Skipped,
                    0, 0, file.SizeBytes, 0, false, null);
                results.Add(result);
                progress?.Report(new ExcelBatchProgress(index + 1, files.Count, file.RelativePath, 1, result));
                continue;
            }

            try
            {
                var outputDirectory = Path.GetDirectoryName(output);
                if (!string.IsNullOrEmpty(outputDirectory)) Directory.CreateDirectory(outputDirectory);
                var inner = new InlineProgress<ExcelOptimizationProgress>(phase => progress?.Report(
                    new ExcelBatchProgress(index, files.Count, file.RelativePath, Fraction(phase), null)));
                var optimized = new ExcelStyleOptimizer().Optimize(file.Path, output, inner, cancellationToken, options);
                result = new ExcelBatchItemResult(file.Path, file.RelativePath, output, ExcelBatchItemStatus.Completed,
                    optimized.InputCellXfCount, optimized.OutputCellXfCount,
                    optimized.InputFileSize, optimized.OutputFileSize, optimized.Verified, null)
                { TableChanges = optimized.TableChanges, AdvancedOptimizationSkipReason = optimized.AdvancedOptimizationSkipReason };
            }
            catch (OperationCanceledException)
            {
                result = new ExcelBatchItemResult(file.Path, file.RelativePath, output, ExcelBatchItemStatus.Cancelled,
                    0, 0, file.SizeBytes, 0, false, null);
                results.Add(result);
                progress?.Report(new ExcelBatchProgress(index + 1, files.Count, file.RelativePath, 1, result));
                return new ExcelBatchOutcome(results, true, files.Count);
            }
            catch (Exception error)
            {
                result = new ExcelBatchItemResult(file.Path, file.RelativePath, output, ExcelBatchItemStatus.Failed,
                    0, 0, file.SizeBytes, 0, false, error.Message);
            }
            results.Add(result);
            progress?.Report(new ExcelBatchProgress(index + 1, files.Count, file.RelativePath, 1, result));
        }
        return new ExcelBatchOutcome(results, false, files.Count);
    }

    private static double Fraction(ExcelOptimizationProgress progress)
    {
        var ratio = progress.Total > 0 ? Math.Clamp((double)progress.Completed / progress.Total, 0, 1) : 0;
        return progress.Phase switch
        {
            ExcelOptimizationPhase.Reading => .05,
            ExcelOptimizationPhase.RewritingWorksheets => .1 + .5 * ratio,
            ExcelOptimizationPhase.RewritingStyles => .65,
            ExcelOptimizationPhase.Verifying => .7 + .2 * ratio,
            ExcelOptimizationPhase.WritingOutput => .98,
            _ => .05
        };
    }

    // Progress<T> posts to the thread pool when no synchronization context exists. Batch callbacks
    // must stay ordered with the per-file completion report, so invoke them inline instead.
    private sealed class InlineProgress<T> : IProgress<T>
    {
        private readonly Action<T> action;
        public InlineProgress(Action<T> action) => this.action = action;
        public void Report(T value) => action(value);
    }
}
