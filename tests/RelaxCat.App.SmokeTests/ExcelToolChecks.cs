using System.IO.Compression;
using System.Text;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using RelaxCat.App;
using RelaxCat.App.Models;

namespace RelaxCat.App.SmokeTests;

/// <summary>Headless checks for the Excel style tool: analysis, worksheet style rendering, optimization and i18n.</summary>
internal static class ExcelToolChecks
{
    internal static void Run(MainWindow window, Action<Task> complete, Action<string> capture, Action<Func<bool>> pumpUntil)
    {
        var theme = window.Workspace.ThemeIndex;
        var language = window.Workspace.LanguageIndex;
        window.Workspace.ThemeIndex = 1; // Light theme shows the real worksheet fill colours clearly.
        window.Workspace.LanguageIndex = 0;
        Dispatcher.UIThread.RunJobs();

        var excelPage = window.FindControl<RelaxCat.App.Views.ExcelToolView>("ExcelToolPage") ??
                        throw new Exception("Missing Excel tool page");
        Click(window.FindControl<Button>("ExcelToolButton")!);
        pumpUntil(() => excelPage.IsVisible && !window.Workspace.HomeOpen);
        Check(window.Workspace.ExcelToolOpen && window.Workspace.SettingsOpen == false, "Excel tool navigation did not open its page");
        capture("excel-tool-empty-light.png");

        var input = Path.GetFullPath("artifacts/excel-tool-sample.xlsx");
        CreateWorkbook(input);
        var output = Path.GetFullPath("artifacts/excel-tool-optimized.xlsx");
        if (File.Exists(output)) File.Delete(output);

        window.Workspace.ExcelTool.FilePath = input;
        Dispatcher.UIThread.RunJobs();
        Check(window.Workspace.ExcelTool.CanAnalyze && !window.Workspace.ExcelTool.CanOptimize,
            "A selected workbook should be analyzable but not yet optimizable");

        complete(window.AnalyzeExcelWorkbookAsync(input));
        pumpUntil(() => window.Workspace.ExcelTool.HasSummary);
        var tool = window.Workspace.ExcelTool;
        Check(tool.Sheets.Count == 2, "Both worksheets were not listed");
        Check(tool.Sheets.Single(sheet => sheet.Name == "Hidden").Hidden, "Hidden worksheet state was not shown");
        var selected = tool.SelectedSheet ?? throw new Exception("No worksheet was selected after analysis");
        Check(selected.Name == "Data", "The first worksheet was not selected");
        Check(selected.Styles.Count == 10, "Worksheet styles were not resolved for the UI");
        Check(selected.Styles.Any(swatch => swatch.Fill is SolidColorBrush), "No filled style swatch was produced");        Check(selected.Styles.Any(swatch => swatch.BorderThickness.Left > 0 || swatch.BorderThickness.Bottom > 0),
            "No bordered style swatch was produced");
        Check(selected.Styles.Any(swatch => swatch.FontWeight == FontWeight.Bold), "Bold font style was not detected");
        Check(selected.Styles[0].Details.Contains("字体"), "Style details are not localized");
        Check(Metric(tool, "可去重格式") == "1,190", "Duplicate format metric is wrong: " + Metric(tool, "可去重格式"));
        Check(Metric(tool, "cellXfs") == "1,200", "cellXfs metric is wrong");
        Check(window.Workspace.ExcelTool.CanOptimize, "Duplicate formats did not enable optimization");
        Check(Metric(tool, "styles.xml ZIP 压缩大小").Length > 0, "Compressed style bytes were not displayed separately");
        Check(tool.Summary!.Diagnostics.Any(row => row.Label == "cellStyleXfs" && row.Value.Contains("已知引用项")),
            "Base style section diagnostics were not localized or displayed");
        Check(excelPage.FindControl<ItemsControl>("SheetTabs")!.GetVisualDescendants().OfType<Button>().Count() == 2,
            "Worksheet tabs were not rendered");
        var swatchCount = excelPage.GetVisualDescendants().OfType<Border>().Count(border => border.Classes.Contains("swatch"));
        Check(swatchCount == 10, "Rendered swatch count was " + swatchCount);
        capture("excel-tool-analysis-light.png");

        complete(window.OptimizeExcelWorkbookAsync(input, output));
        pumpUntil(() => tool.HasResult);
        Check(File.Exists(output), "Optimized workbook was not written");
        Check(tool.Result!.Verified && tool.Result.Checks.Count == 5 && tool.Result.Checks.All(check => check.Passed),
            "Optimization verification did not pass");
        Check(tool.FilePath == Path.GetFullPath(output), "The page did not move to the optimized workbook");
        Check(Metric(tool, "cellXfs") == "10", "The optimized workbook should keep ten unique formats");
        Check(!tool.CanOptimize, "An optimized workbook still offered optimization");
        Check(tool.Result.Metrics.Any(metric => metric.Label == "cellXfs" && metric.Value.Contains("10")),
            "Optimization result did not report the new cell format count");
        Check(tool.Result.OutputName == "excel-tool-optimized.xlsx", "Output workbook name was not shown");
        Check(tool.SelectedSheet!.Styles.Count == 10, "The optimized workbook did not expose its styles");
        capture("excel-tool-result-light.png");

        // Minimum window size: both panels stay reachable and the primary actions stay visible.
        window.Width = 1000; window.Height = 680; Dispatcher.UIThread.RunJobs();
        Check(excelPage.FindControl<Button>("OptimizeButton")!.IsEffectivelyVisible,
            "Excel tool actions disappeared at the minimum window size");
        capture("excel-tool-narrow-light.png");
        window.Width = 1280; window.Height = 850; Dispatcher.UIThread.RunJobs();

        window.Workspace.ThemeIndex = 0;
        Dispatcher.UIThread.RunJobs();
        capture("excel-tool-result-dark.png");

        window.Workspace.LanguageIndex = 1;
        Dispatcher.UIThread.RunJobs();
        Check(Metric(tool, "ファイルサイズ").Length > 0, "Japanese metric labels did not refresh");
        Check(tool.Summary!.Diagnostics.Any(row => row.Label == "cellStyleXfs" && row.Value.Contains("既知の参照項目")),
            "Japanese diagnostics did not refresh");
        Check(tool.SelectedSheet!.Styles[0].Details.Contains("フォント"), "Japanese style details did not refresh");
        Check(Label(excelPage.FindControl<Button>("OptimizeButton")!).StartsWith("最適化"), "Japanese action label missing");
        window.Workspace.ThemeIndex = 1;
        Dispatcher.UIThread.RunJobs();
        capture("excel-tool-result-ja.png");

        window.Workspace.LanguageIndex = 0;
        window.Workspace.ThemeIndex = 1;
        Dispatcher.UIThread.RunJobs();

        // Same-file output is refused without touching either workbook.
        var before = File.ReadAllBytes(input);
        complete(window.OptimizeExcelWorkbookAsync(input, input));
        Check(window.Workspace.Message == window.Workspace.Text["ExcelSameFile"], "Same-file output was not rejected");
        Check(before.SequenceEqual(File.ReadAllBytes(input)), "Rejected optimization changed the input workbook");

        // Broken files report a localized error and clear stale analysis.
        var broken = Path.GetFullPath("artifacts/excel-tool-broken.xlsx");
        File.WriteAllText(broken, "not a zip package");
        complete(window.AnalyzeExcelWorkbookAsync(broken));
        Check(window.Workspace.Message.StartsWith(window.Workspace.Text["ExcelAnalysisFailed"]),
            "Broken workbook did not report an analysis error");
        Check(!window.Workspace.ExcelTool.HasSummary && !window.Workspace.ExcelTool.HasResult,
            "Broken workbook left stale analysis visible");

        complete(window.AnalyzeExcelWorkbookAsync(input));
        pumpUntil(() => window.Workspace.ExcelTool.HasSummary);
        // GC changes invalidate the projection; analysis and optimization must use the same option.
        var gcOption = excelPage.FindControl<CheckBox>("CollectUnusedStylesOption")!;
        gcOption.IsChecked = true;
        Dispatcher.UIThread.RunJobs();
        Check(tool.CollectUnusedStyles && !tool.HasSummary && !tool.CanOptimize,
            "GC toggle left an obsolete projection usable");
        complete(window.AnalyzeExcelWorkbookAsync(input));
        Check(tool.HasSummary && tool.CanOptimize, "GC analysis did not restore the actions");
        var gcOutput = Path.GetFullPath("artifacts/excel-tool-gc.xlsx");
        complete(window.OptimizeExcelWorkbookAsync(input, gcOutput));
        Check(tool.HasResult && tool.Result!.Verified, "GC was not passed through the desktop workflow");
        capture("excel-tool-gc-light.png");
        gcOption.IsChecked = false;
        Dispatcher.UIThread.RunJobs();
        complete(window.AnalyzeExcelWorkbookAsync(input));
        Click(window.FindControl<Button>("HomeButton")!);
        window.Workspace.LanguageIndex = language;
        window.Workspace.ThemeIndex = theme;
        Dispatcher.UIThread.RunJobs();
        Console.WriteLine("PASS: Excel style analysis, worksheet style previews, verified deduplication, same-file refusal and bilingual refresh.");
    }

    private static string Metric(ExcelToolModel tool, string label) =>
        tool.Summary!.Metrics.FirstOrDefault(metric => metric.Label == label)?.Value ?? "<missing:" + label + ">";

    private static string Label(Button button) => button.Content switch
    {
        string text => text,
        Panel panel => string.Concat(panel.Children.OfType<TextBlock>().Select(text => text.Text)),
        _ => ""
    };

    private static void Click(Button button)
    {
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        Dispatcher.UIThread.RunJobs();
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }

    /// <summary>Creates a workbook with 1,200 cell formats of which only ten are unique.</summary>
    internal static void CreateWorkbook(string path)
    {
        if (File.Exists(path)) File.Delete(path);
        var styles = BuildStyles();
        var data = new StringBuilder();
        for (var row = 1; row <= 30; row++)
        {
            data.Append($"<row r=\"{row}\">");
            for (var column = 1; column <= 6; column++)
            {
                var reference = (char)('A' + column - 1) + row.ToString();
                var style = (row * 6 + column) % 10;
                data.Append(style == 0
                    ? $"<c r=\"{reference}\"><v>{row * column}</v></c>"
                    : $"<c r=\"{reference}\" s=\"{style}\"><v>{row * column}</v></c>");
            }
            data.Append("</row>");
        }
        data.Append("<row r=\"31\" s=\"5\"><c r=\"A31\"><v>1</v></c><c r=\"B31\"><v>2</v></c></row>");
        var sheet = Worksheet("A1:F31", data.ToString());
        var hidden = Worksheet("A1", """<row r="1"><c r="A1" s="7"><v>1</v></c></row>""");

        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes());
        Add(archive, "_rels/.rels", RootRelationships());
        Add(archive, "xl/workbook.xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
                <sheet name="Data" sheetId="1" r:id="rId1"/>
                <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
              </sheets>
            </workbook>
            """);
        Add(archive, "xl/_rels/workbook.xml.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            </Relationships>
            """);
        Add(archive, "xl/styles.xml", styles);
        Add(archive, "xl/worksheets/sheet1.xml", sheet);
        Add(archive, "xl/worksheets/sheet2.xml", hidden);
    }

    private static string BuildStyles()
    {
        const string fonts = """
            <fonts count="4">
              <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
              <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>
              <font><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
              <font><i/><sz val="12"/><color rgb="FF1F3864"/><name val="Times New Roman"/><family val="1"/></font>
            </fonts>
            """;
        const string fills = """
            <fills count="7">
              <fill><patternFill patternType="none"/></fill>
              <fill><patternFill patternType="gray125"/></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill>
              <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
            </fills>
            """;
        const string borders = """
            <borders count="3">
              <border><left/><right/><top/><bottom/><diagonal/></border>
              <border><left/><right/><top/><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>
              <border><left style="medium"><color rgb="FF2F5597"/></left><right style="medium"><color rgb="FF2F5597"/></right><top style="medium"><color rgb="FF2F5597"/></top><bottom style="medium"><color rgb="FF2F5597"/></bottom><diagonal/></border>
            </borders>
            """;
        const string formats = """
            <numFmts count="2">
              <numFmt numFmtId="164" formatCode="0.00%"/>
              <numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
            </numFmts>
            """;
        var unique = new[]
        {
            """<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>""",
            """<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>""",
            """<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>""",
            """<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>""",
            """<xf numFmtId="164" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>""",
            """<xf numFmtId="0" fontId="3" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>""",
            """<xf numFmtId="0" fontId="3" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1"/>""",
            """<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>""",
            """<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>""",
            """<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>"""
        };
        var cellXfs = new StringBuilder();
        for (var index = 0; index < 1200; index++)
            cellXfs.Append(unique[index % unique.Length]);
        return $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            {formats}
            {fonts}
            {fills}
            {borders}
            <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
            <cellXfs count="1200">{cellXfs}</cellXfs>
            <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
            <dxfs count="0"/>
            <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
            </styleSheet>
            """;
    }

    private static string Worksheet(string dimension, string sheetData) => $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="{dimension}"/>
          <sheetViews><sheetView workbookViewId="0"/></sheetViews>
          <sheetFormatPr defaultRowHeight="15"/>
          <sheetData>{sheetData}</sheetData>
          <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        </worksheet>
        """;

    private static string ContentTypes() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private static string RootRelationships() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static void Add(ZipArchive archive, string name, string contents)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(contents);
    }
}
