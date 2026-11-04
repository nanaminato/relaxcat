using System.IO.Compression;
using System.Text;
using RelaxCat.App.Models;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.SmokeTests;

internal static class ExtensionDiagnosticsChecks
{
    internal static void Run()
    {
        var path = Path.GetFullPath("artifacts/extension-diagnostics-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            CreateWorkbook(path);
            var text = new UiText();
            var model = new ExcelToolModel(text) { FilePath = path };
            var analysis = new ExcelStyleAnalyzer().Analyze(path, options: new ExcelStyleOptimizationOptions(
                CollectUnusedStyles: true, RemoveCopiedNamedStyles: true));
            model.ApplyAnalysis(analysis);
            Check(analysis.AdvancedOptimizationSkipReason is not null && !model.CanOptimize,
                "Extension diagnostics unexpectedly enabled optimization of unknown extensions");
            Check(analysis.ExtensionDiagnostics.Entries.Count > 0, "No extension types reached the presentation model");
            foreach (var language in text.AvailableLanguages)
            {
                text.CurrentCulture = language.Culture;
                model.RefreshLanguage();
                var rows = ExtensionRows(model, text);
                Check(rows.Any(row => row.Label == text["ExtensionDiagnosticsTitle"] && row.Value == text["ExtensionDiagnosticsHint"]),
                    "Local extension metadata warning did not refresh: " + language.Culture);
                Check(rows.Any(row => row.Label == text["ExtensionDiagnosticsScope"] && row.Value == text["ExtensionDiagnosticsScopeHint"]),
                    "Extension diagnostic scope did not refresh: " + language.Culture);
                Check(rows.Any(row => row.Label.StartsWith(text["ExtensionDiagnosticUri"]) && row.Value == "urn:relaxcat:diagnostic-test"),
                    "Extension URI type was not shown: " + language.Culture);
                foreach (var entry in analysis.ExtensionDiagnostics.Entries)
                    Check(rows.Any(row => row.Value == entry.Name), "Collected type missing from diagnostics: " + entry.Kind);
                var display = string.Join("\n", rows.Select(row => row.Label + " " + row.Value));
                Check(!display.Contains("PRIVATE_ATTRIBUTE_VALUE") && !display.Contains("PRIVATE_EXTENSION_TEXT") &&
                      !display.Contains(path) && !display.Contains("Normal"),
                    "Extension diagnostic display included payload text, paths or style names");
                Check(!model.CanOptimize, "Changing diagnostic language relaxed the unknown-extension block");
            }

            var bounded = Enumerable.Range(0, 64).Select(index =>
                new ExtensionTypeEntry("element", new string('x', 125) + index.ToString("D3"), index + 1)).ToArray();
            model.ApplyAnalysis(analysis with { ExtensionDiagnostics = new(bounded, true) });
            var boundedRows = ExtensionRows(model, text);
            Check(boundedRows.Any(row => row.Label == text["ExtensionDiagnosticsLimit"] && row.Value == text["ExtensionDiagnosticsTruncated"]),
                "Truncated extension metadata did not warn that counts may be partial");
            Check(boundedRows.Count(row => row.Label.StartsWith(text["ExtensionDiagnosticElement"])) == 64 &&
                  boundedRows.Count(row => row.Value.Length == 128) >= 64,
                "Bounded extension metadata did not preserve all displayed type entries");
            model.ApplyAnalysis(analysis with { ExtensionDiagnostics = new([], false) });
            Check(ExtensionRows(model, text).Any(row => row.Value == text["ExtensionDiagnosticsEmpty"]),
                "Empty diagnostic scope was not distinguished from complete package coverage");
            Console.WriteLine("PASS: local extension diagnostics, four-language privacy/scope warnings, partial counts and preserved optimization block.");
        }
        finally { File.Delete(path); }
    }

    private static ExcelMetric[] ExtensionRows(ExcelToolModel model, UiText text) =>
        model.Summary!.Diagnostics.SkipWhile(row => row.Label != text["ExtensionDiagnosticsTitle"]).ToArray();

    internal static void CreateWorkbook(string path)
    {
        ExcelToolChecks.CreateWorkbook(path);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry("xl/styles.xml")!;
        string styles;
        using (var reader = new StreamReader(entry.Open())) styles = reader.ReadToEnd();
        styles = styles.Replace("</styleSheet>", """
            <extLst><ext uri="urn:relaxcat:diagnostic-test"><custom:metadata xmlns:custom="urn:relaxcat:custom-test"
            secret="PRIVATE_ATTRIBUTE_VALUE">PRIVATE_EXTENSION_TEXT</custom:metadata></ext></extLst></styleSheet>
            """);
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry("xl/styles.xml").Open(), new UTF8Encoding(false));
        writer.Write(styles);
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
