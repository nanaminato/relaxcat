using System.Globalization;
using System.Text.Json;
using RelaxCat.Core.Excel;

return Run(args);

static int Run(string[] args)
{
    if (args.Length == 0 || args[0] is "-h" or "--help" or "help")
    {
        PrintUsage();
        return args.Length == 0 ? 2 : 0;
    }
    var command = args[0].ToLowerInvariant();
    try
    {
        return command switch
        {
            "analyze" => Analyze(args[1..]),
            "optimize" => Optimize(args[1..]),
            _ => Unknown(command)
        };
    }
    catch (UnauthorizedAccessException error)
    {
        Console.Error.WriteLine("error: access denied. Choose a writable output folder or trust this tool in your security software. " + error.Message);
        return 1;
    }
    catch (Exception error)
    {
        Console.Error.WriteLine("error: " + error.Message);
        return 1;
    }
}

static int Analyze(string[] args)
{
    var path = Positional(args);
    var json = args.Contains("--json");
    var analysis = new ExcelStyleAnalyzer().Analyze(path);
    if (json)
    {
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            analysis.Path,
            analysis.FileSize,
            analysis.MacroEnabled,
            analysis.StylesXmlBytes,
            analysis.ProjectedStylesXmlBytes,
            analysis.FontCount,
            analysis.FillCount,
            analysis.BorderCount,
            analysis.CellStyleXfCount,
            analysis.CellXfCount,
            analysis.UniqueCellXfCount,
            analysis.DuplicateCellXfCount,
            analysis.ReferencedCellXfCount,
            analysis.UnreferencedCellXfCount,
            analysis.ReductionRatio,
            Worksheets = analysis.Worksheets.Select(sheet => new
            {
                sheet.Name,
                sheet.Hidden,
                sheet.UsedRange,
                sheet.CellCount,
                sheet.DistinctStyleCount
            })
        }, new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }

    Console.WriteLine("Workbook:   " + analysis.Path);
    Console.WriteLine("File size:  " + Bytes(analysis.FileSize));
    Console.WriteLine("styles.xml: " + Bytes(analysis.StylesXmlBytes) + "  (projected " + Bytes(analysis.ProjectedStylesXmlBytes) + ")");
    Console.WriteLine("Macros:     " + (analysis.MacroEnabled ? "yes" : "no"));
    Console.WriteLine();
    Console.WriteLine($"Fonts: {Count(analysis.FontCount)}   Fills: {Count(analysis.FillCount)}   Borders: {Count(analysis.BorderCount)}   cellStyleXfs: {Count(analysis.CellStyleXfCount)}");
    Console.WriteLine();
    Console.WriteLine("cellXfs total:       " + Count(analysis.CellXfCount).PadLeft(10));
    Console.WriteLine("Unique formats:      " + Count(analysis.UniqueCellXfCount).PadLeft(10));
    Console.WriteLine("Duplicate formats:   " + Count(analysis.DuplicateCellXfCount).PadLeft(10));
    Console.WriteLine("Referenced (sheets): " + Count(analysis.ReferencedCellXfCount).PadLeft(10));
    Console.WriteLine("Unreferenced:        " + Count(analysis.UnreferencedCellXfCount).PadLeft(10));
    Console.WriteLine("Reducible:           " + Percent(analysis.ReductionRatio).PadLeft(10));
    Console.WriteLine();
    if (analysis.Worksheets.Count == 0) Console.WriteLine("No worksheets.");
    foreach (var sheet in analysis.Worksheets)
    {
        var hidden = sheet.Hidden ? " [hidden]" : "";
        Console.WriteLine($"  {sheet.Name}{hidden}  {sheet.UsedRange,-12}  cells {Count(sheet.CellCount),-8}  styles {Count(sheet.DistinctStyleCount)}");
    }
    return 0;
}

static int Optimize(string[] args)
{
    string? output = null;
    var positional = new List<string>();
    for (var index = 0; index < args.Length; index++)
    {
        var argument = args[index];
        if (argument is "--output" or "-o")
        {
            if (++index >= args.Length) throw new ArgumentException("--output requires a path.");
            output = args[index];
        }
        else if (argument.StartsWith('-')) throw new ArgumentException("Unknown option " + argument);
        else positional.Add(argument);
    }
    if (positional.Count != 1) throw new ArgumentException("optimize requires exactly one input workbook.");
    if (output is null) throw new ArgumentException("optimize requires --output <path>.");
    var result = new ExcelStyleOptimizer().Optimize(positional[0], output);
    Console.WriteLine("Output: " + result.OutputPath);
    Console.WriteLine("File size:  " + Bytes(result.InputFileSize) + " -> " + Bytes(result.OutputFileSize) + "  (" + Percent(result.ReductionRatio) + " fewer cell formats)");
    Console.WriteLine("cellXfs:    " + Count(result.InputCellXfCount) + " -> " + Count(result.OutputCellXfCount));
    Console.WriteLine("styles.xml: " + Bytes(result.InputStylesXmlBytes) + " -> " + Bytes(result.OutputStylesXmlBytes));
    Console.WriteLine("References remapped: " + Count(result.RemappedReferences) + " across " + Count(result.WorksheetCount) + " worksheet(s)");
    Console.WriteLine();
    Console.WriteLine("Verification: " + (result.Verified ? "passed" : "FAILED"));
    foreach (var check in result.Checks)
        Console.WriteLine("  " + (check.Passed ? "ok  " : "FAIL") + "  " + check.Key + "  " + check.Detail);
    return result.Verified ? 0 : 1;
}

static string Positional(string[] args)
{
    var path = args.FirstOrDefault(argument => !argument.StartsWith('-'))
               ?? throw new ArgumentException("A workbook path is required.");
    return path;
}

static int Unknown(string command)
{
    Console.Error.WriteLine("error: unknown command '" + command + "'");
    PrintUsage();
    return 2;
}

static void PrintUsage()
{
    Console.WriteLine("""
        RelaxCat Excel style tool

        Usage:
          RelaxCat.ExcelTool analyze <workbook.xlsx|workbook.xlsm> [--json]
          RelaxCat.ExcelTool optimize <workbook.xlsx|workbook.xlsm> --output <copy.xlsx|copy.xlsm>

        analyze   Read-only report of the style table, duplicate cell formats and worksheet styles.
        optimize  Merge XML-identical cell formats into a NEW workbook. The input is never modified.
                  The output keeps the input's extension; .xlsm files keep their VBA project and all
                  other package parts byte-for-byte. Every result is verified before it is published.

        Exit codes: 0 success, 1 failure, 2 usage error.
        """);
}

static string Count(long value) => value.ToString("N0", CultureInfo.InvariantCulture);

static string Percent(double ratio) => (ratio * 100).ToString("0.##", CultureInfo.InvariantCulture) + "%";

static string Bytes(long value) => value switch
{
    >= 1024L * 1024 * 1024 => (value / (1024d * 1024 * 1024)).ToString("0.##", CultureInfo.InvariantCulture) + " GB",
    >= 1024L * 1024 => (value / (1024d * 1024)).ToString("0.##", CultureInfo.InvariantCulture) + " MB",
    >= 1024 => (value / 1024d).ToString("0.#", CultureInfo.InvariantCulture) + " KB",
    _ => value.ToString("N0", CultureInfo.InvariantCulture) + " B"
};
