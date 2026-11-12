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
                var rows = model.Summary!.Diagnostics;
                Check(text["PrivateDiagnosticsTitle"] != "PrivateDiagnosticsTitle" &&
                      text["PrivateDiagnosticsHint"] != "PrivateDiagnosticsHint",
                    "Private diagnostic localization missing: " + language.Culture);
                foreach (var reason in analysis.AdvancedOptimizationSkipReasons)
                    Check(rows.Any(row => row.Value == text["AdvancedSkip_" + reason]),
                        "Detected blocker missing after language refresh: " + reason);
                var report = model.PrivateDiagnosticJson;
                Check(report == PrivateStyleDiagnostics.FromAnalysis(analysis).ToJson(),
                    "Private diagnostic projection drifted after language refresh");
                var display = report + "\n" + string.Join("\n", rows.Select(row => row.Label + " " + row.Value));
                foreach (var secret in new[] { "PRIVATE_ATTRIBUTE_VALUE", "PRIVATE_EXTENSION_TEXT", path,
                    "urn:relaxcat:diagnostic-test", "urn:relaxcat:custom-test" })
                    Check(!display.Contains(secret), "Private diagnostics included raw metadata or payload");
                Check(!model.CanOptimize, "Changing diagnostic language relaxed the unknown-extension block");
            }

            var bounded = Enumerable.Range(0, 64).Select(index =>
                new ExtensionTypeEntry("element", new string('x', 125) + index.ToString("D3"), index + 1)).ToArray();
            model.ApplyAnalysis(analysis with { ExtensionDiagnostics = new(bounded, true) });
            Check(model.PrivateDiagnosticJson == PrivateStyleDiagnostics.FromAnalysis(analysis).ToJson(),
                "Arbitrary extension metadata leaked into the private report");
            Check(!model.Summary!.Diagnostics.Any(row => row.Value.Contains(new string('x', 125))),
                "Raw extension names were displayed by default");
            model.ApplyAnalysis(analysis with { ExtensionDiagnostics = new([], false) });
            Check(model.PrivateDiagnosticJson == PrivateStyleDiagnostics.FromAnalysis(analysis).ToJson(),
                "Empty raw metadata incorrectly changed gate diagnostics");
            model.Clear();
            Check(model.PrivateDiagnosticJson.Length == 0 && !model.HasSummary,
                "Clearing the model retained a stale diagnostic report");
            model.ApplyAnalysis(analysis);
            model.DestructiveStyleCleanup = true;
            Check(model.PrivateDiagnosticJson.Length == 0 && !model.HasSummary,
                "Changing mode retained diagnostics for the previous policy");
            model.ApplyAnalysis(analysis);
            model.FilePath = path + ".different.xlsx";
            Check(model.PrivateDiagnosticJson.Length == 0 && !model.HasSummary,
                "Changing input retained another workbook's diagnostics");
            Console.WriteLine("PASS: private diagnostics, four-language blockers, raw metadata exclusion and stale-report invalidation.");
        }
        finally { File.Delete(path); }
    }

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
