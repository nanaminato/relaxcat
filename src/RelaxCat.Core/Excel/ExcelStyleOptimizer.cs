using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace RelaxCat.Core.Excel;

/// <summary>
/// Rewrites only the parts needed to merge provably identical cell formats. Every other
/// package part is copied byte-for-byte, and the result is verified before it is published.
/// </summary>
public sealed class ExcelStyleOptimizer
{
    public ExcelOptimizationResult Optimize(string inputPath, string outputPath,
        IProgress<ExcelOptimizationProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("A workbook path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        var input = Path.GetFullPath(inputPath);
        var output = Path.GetFullPath(outputPath);
        if (!File.Exists(input)) throw new FileNotFoundException("The workbook was not found.", input);
        if (!ExcelStyleAnalyzer.IsSupported(input))
            throw new NotSupportedException("Only .xlsx and .xlsm workbooks can be optimized; .xls is a different binary format.");
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The optimized copy must be written to a different file; the original workbook is never overwritten.");
        if (!string.Equals(Path.GetExtension(input), Path.GetExtension(output), StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The optimized copy must keep the input workbook's extension so its package type stays valid.");
        if (Directory.Exists(output)) throw new ArgumentException("The output path is a folder, not a workbook.", nameof(outputPath));
        var directory = Path.GetDirectoryName(output);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
            throw new DirectoryNotFoundException("The output folder does not exist.");
        var temporary = Path.Combine(directory,
            "." + Path.GetFileNameWithoutExtension(output) + "." + Guid.NewGuid().ToString("N") + ".partial");

        try
        {
            using var package = new OoxmlPackage(input);
            var styles = StyleSheetDocument.Load(package);
            var worksheets = ExcelStyleAnalyzer.ReadWorksheets(package);
            var inputFileSize = new FileInfo(input).Length;
            var inputStylesBytes = package.LengthOf("xl/styles.xml");
            progress?.Report(new ExcelOptimizationProgress(ExcelOptimizationPhase.Reading, 0, worksheets.Count));
            cancellationToken.ThrowIfCancellationRequested();
            var plan = StyleDeduplicationPlan.Build(styles);
            var inputKeys = plan.OriginalKeys;
            var outputKeys = plan.Kept.Select(OoxmlXml.CanonicalKey).ToArray();
            var hasDuplicates = plan.KeptCount != styles.CellXfs.Count;

            var changed = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            if (hasDuplicates) changed["xl/styles.xml"] = plan.StylesXml;
            var remapped = 0;
            for (var index = 0; index < worksheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var sheet = worksheets[index];
                progress?.Report(new ExcelOptimizationProgress(ExcelOptimizationPhase.RewritingWorksheets, index, worksheets.Count));
                if (!hasDuplicates) continue;
                var document = package.ReadXml(sheet.PartName);
                remapped += RemapWorksheet(document, plan.Map, styles.CellXfs.Count, sheet.Name);
                changed[sheet.PartName] = OoxmlXml.Serialize(document);
            }
            progress?.Report(new ExcelOptimizationProgress(ExcelOptimizationPhase.RewritingStyles, worksheets.Count, worksheets.Count));

            using (var outputStream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create))
            {
                foreach (var entry in package.Archive.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (changed.TryGetValue(entry.FullName, out var bytes)) WriteEntry(outputArchive, entry, bytes);
                    else CopyEntry(outputArchive, entry);
                }
            }

            progress?.Report(new ExcelOptimizationProgress(ExcelOptimizationPhase.Verifying, 0, worksheets.Count));
            long outputStylesBytes;
            var checks = new List<OptimizationCheck>();
            using (var outputPackage = new OoxmlPackage(temporary))
            {
                var outputStyles = StyleSheetDocument.Load(outputPackage);
                outputStylesBytes = outputPackage.LengthOf("xl/styles.xml");
                checks.Add(CheckStyleSections(styles.Document, outputStyles.Document));
                checks.Add(CheckCellXfs(plan, outputStyles));
                checks.Add(CheckWorksheetReferences(package, outputPackage, worksheets, inputKeys, outputKeys));
                checks.Add(CheckPreservedParts(package, outputPackage, changed));
            }
            checks.Add(CheckOutputFile(temporary));
            var verified = checks.All(check => check.Passed);
            progress?.Report(new ExcelOptimizationProgress(ExcelOptimizationPhase.WritingOutput, 1, 1));
            if (!verified)
                throw new InvalidDataException(
                    "The optimized workbook failed verification and was discarded: " +
                    string.Join("; ", checks.Where(check => !check.Passed).Select(check => $"{check.Key} ({check.Detail})")));

            File.Move(temporary, output, overwrite: true);
            return new ExcelOptimizationResult(input, output, inputFileSize, new FileInfo(output).Length,
                styles.CellXfs.Count, plan.KeptCount, inputStylesBytes, outputStylesBytes,
                remapped, worksheets.Count, checks, true);
        }
        finally
        {
            TryDelete(temporary);
        }
    }

    private static int RemapWorksheet(XDocument document, int[] map, int cellFormatCount, string sheetName)
    {
        var remapped = 0;
        foreach (var element in document.Descendants())
        {
            var attributeName = element.Name.LocalName switch
            {
                "c" or "row" => "s",
                "col" => "style",
                _ => null
            };
            if (attributeName is null) continue;
            var attribute = element.Attributes().FirstOrDefault(candidate =>
                candidate.Name.LocalName == attributeName && !candidate.IsNamespaceDeclaration);
            if (attribute is null) continue;
            if (!int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var style) ||
                style < 0 || style >= cellFormatCount)
                throw new InvalidDataException(
                    $"Worksheet '{sheetName}' references cell format '{attribute.Value}', which the workbook does not define.");
            var replacement = map[style];
            if (replacement == style) continue;
            attribute.Value = replacement.ToString(CultureInfo.InvariantCulture);
            remapped++;
        }
        return remapped;
    }

    private static OptimizationCheck CheckStyleSections(XDocument input, XDocument output)
    {
        var inputSections = Sections(input);
        var outputSections = Sections(output);
        if (inputSections.Count != outputSections.Count)
            return new OptimizationCheck("stylesSections", false, $"{inputSections.Count} → {outputSections.Count}");
        foreach (var (name, key) in inputSections)
            if (!outputSections.TryGetValue(name, out var other) || other != key)
                return new OptimizationCheck("stylesSections", false, name);
        return new OptimizationCheck("stylesSections", true, inputSections.Count.ToString(CultureInfo.InvariantCulture));

        static Dictionary<string, string> Sections(XDocument document) =>
            (document.Root?.Elements() ?? Enumerable.Empty<XElement>())
            .Where(element => element.Name.LocalName != "cellXfs")
            .ToDictionary(element => element.Name.LocalName, OoxmlXml.CanonicalKey);
    }

    private static OptimizationCheck CheckCellXfs(StyleDeduplicationPlan plan, StyleSheetDocument output)
    {
        if (output.CellXfs.Count != plan.KeptCount)
            return new OptimizationCheck("cellXfs", false, $"{plan.KeptCount} → {output.CellXfs.Count}");
        for (var index = 0; index < plan.KeptCount; index++)
            if (OoxmlXml.CanonicalKey(output.CellXfs[index]) != OoxmlXml.CanonicalKey(plan.Kept[index]))
                return new OptimizationCheck("cellXfs", false, "index " + index.ToString(CultureInfo.InvariantCulture));
        var declared = OoxmlXml.IntAttribute(OoxmlXml.Child(output.Root, "cellXfs"), "count");
        return declared == plan.KeptCount
            ? new OptimizationCheck("cellXfs", true, plan.KeptCount.ToString(CultureInfo.InvariantCulture))
            : new OptimizationCheck("cellXfs", false, $"count attribute is {declared?.ToString(CultureInfo.InvariantCulture) ?? "missing"}");
    }

    private static OptimizationCheck CheckWorksheetReferences(OoxmlPackage input, OoxmlPackage output,
        IReadOnlyList<WorkbookSheet> worksheets, IReadOnlyList<string> inputKeys, IReadOnlyList<string> outputKeys)
    {
        var cells = 0;
        foreach (var sheet in worksheets)
        {
            var before = WorksheetStyleScanner.Scan(input.ReadXml(sheet.PartName), inputKeys.Count, sheet.Name);
            var after = WorksheetStyleScanner.Scan(output.ReadXml(sheet.PartName), outputKeys.Count, sheet.Name);
            if (before.CellCount != after.CellCount || before.UsedRange != after.UsedRange)
                return new OptimizationCheck("references", false, sheet.Name);
            if (!SameUsage(before.EffectiveStyleCounts, after.EffectiveStyleCounts, inputKeys, outputKeys) ||
                !SameUsage(before.RawStyleCounts, after.RawStyleCounts, inputKeys, outputKeys))
                return new OptimizationCheck("references", false, sheet.Name + " styles");
            cells += after.CellCount;
        }
        return new OptimizationCheck("references", true, cells.ToString("N0", CultureInfo.InvariantCulture));

        static bool SameUsage(Dictionary<int, int> before, Dictionary<int, int> after,
            IReadOnlyList<string> inputKeys, IReadOnlyList<string> outputKeys)
        {
            var left = Aggregate(before, inputKeys);
            var right = Aggregate(after, outputKeys);
            return left.Count == right.Count && left.All(pair => right.GetValueOrDefault(pair.Key) == pair.Value);

            static Dictionary<string, int> Aggregate(Dictionary<int, int> counts, IReadOnlyList<string> keys)
            {
                var result = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var (style, count) in counts)
                {
                    if (style < 0 || style >= keys.Count) return new Dictionary<string, int>(StringComparer.Ordinal);
                    var key = keys[style];
                    result[key] = result.GetValueOrDefault(key) + count;
                }
                return result;
            }
        }
    }

    private static OptimizationCheck CheckPreservedParts(OoxmlPackage input, OoxmlPackage output,
        IReadOnlyDictionary<string, byte[]> changed)
    {
        var outputParts = output.Archive.Entries.ToDictionary(entry => entry.FullName, StringComparer.OrdinalIgnoreCase);
        if (outputParts.Count != input.Archive.Entries.Count)
            return new OptimizationCheck("parts", false, $"{input.Archive.Entries.Count} → {outputParts.Count}");
        var preserved = 0;
        foreach (var entry in input.Archive.Entries)
        {
            if (!outputParts.TryGetValue(entry.FullName, out var other))
                return new OptimizationCheck("parts", false, "missing " + entry.FullName);
            if (changed.ContainsKey(entry.FullName)) continue;
            if (other.Length != entry.Length)
                return new OptimizationCheck("parts", false, entry.FullName);
            preserved++;
        }
        return new OptimizationCheck("parts", true, preserved.ToString("N0", CultureInfo.InvariantCulture));
    }

    private static OptimizationCheck CheckOutputFile(string path) =>
        File.Exists(path) && new FileInfo(path).Length > 0
            ? new OptimizationCheck("package", true, "")
            : new OptimizationCheck("package", false, "missing output");

    private static void WriteEntry(ZipArchive archive, ZipArchiveEntry original, byte[] bytes)
    {
        var entry = archive.CreateEntry(original.FullName, CompressionLevel.Optimal);
        CopyMetadata(original, entry);
        using var target = entry.Open();
        target.Write(bytes);
    }

    private static void CopyEntry(ZipArchive archive, ZipArchiveEntry original)
    {
        var entry = archive.CreateEntry(original.FullName, CompressionLevel.Optimal);
        CopyMetadata(original, entry);
        using var source = original.Open();
        using var target = entry.Open();
        source.CopyTo(target);
    }

    private static void CopyMetadata(ZipArchiveEntry source, ZipArchiveEntry target)
    {
        try { target.LastWriteTime = source.LastWriteTime; }
        catch (ArgumentOutOfRangeException) { }
        target.ExternalAttributes = source.ExternalAttributes;
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
