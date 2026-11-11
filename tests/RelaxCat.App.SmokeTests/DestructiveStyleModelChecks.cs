using System.IO.Compression;
using System.Xml.Linq;
using RelaxCat.App.Models;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.SmokeTests;

internal static class DestructiveStyleModelChecks
{
    internal static void Run()
    {
        var directory = Path.GetFullPath("artifacts/destructive-model-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var input = Path.Combine(directory, "input.xlsx");
        var output = Path.Combine(directory, "output.xlsx");
        try
        {
            ExcelToolChecks.CreateNamedStyleWorkbook(input);
            var original = File.ReadAllBytes(input);
            var text = new UiText();
            var model = new ExcelToolModel(text) { FilePath = input };
            Check(!model.DestructiveStyleCleanup && !model.Batch.DestructiveStyleCleanup, "Destructive defaults must be off");
            model.ApplyAnalysis(new ExcelStyleAnalyzer().Analyze(input));
            model.DestructiveStyleCleanup = true;
            Check(!model.HasSummary && !model.CanOptimize && model.OptimizationOptions.DestructiveStyleCleanup &&
                !model.RemoveCopiedNamedStyles, "Destructive toggle did not invalidate the analysis independently");
            model.ApplyAnalysis(new ExcelStyleAnalyzer().Analyze(input, options: model.OptimizationOptions));
            Check(model.CanOptimize, "Destructive analysis did not allow named-style flattening");
            var result = new ExcelStyleOptimizer().Optimize(input, output, options: model.OptimizationOptions);
            model.ApplyOptimization(result);
            foreach (var language in text.AvailableLanguages)
            {
                text.CurrentCulture = language.Culture;
                model.RefreshLanguage();
                Check(text["DestructiveStyleCleanup"] != "DestructiveStyleCleanup" &&
                    text["DestructiveStyleCleanupHint"] != "DestructiveStyleCleanupHint", "Missing localized risk warning");
                Check(model.Result!.VerificationText == text["DestructiveVerification"], "Output falsely implies visual verification");
                Check(model.Summary!.Diagnostics.Any(row => row.Label == text["DestructiveLoss"]), "Analysis omitted loss label");
            }
            Check(result.IsLossy && result.Verified && original.SequenceEqual(File.ReadAllBytes(input)), "Source or output invariant failed");
            model.DestructiveStyleCleanup = false;
            Check(!model.HasResult && !model.HasSummary, "Disabling destructive mode left stale results");

            var batch = model.Batch;
            var destination = Path.Combine(directory, "batch");
            batch.InputFolder = directory;
            batch.OutputFolder = destination;
            batch.ApplyDiscovery(ExcelWorkbookFinder.Find(directory, false, destination));
            batch.DestructiveStyleCleanup = true;
            var outcome = new ExcelBatchOptimizer().Optimize(batch.Files, destination, overwrite: false, options: batch.OptimizationOptions);
            foreach (var item in outcome.Results) batch.ApplyResult(item);
            Check(outcome.Results.All(item => item.Status == ExcelBatchItemStatus.Completed && item.IsLossy), "Batch did not pass destructive option");
            Check(batch.Items.All(item => item.DetailText.Contains(text["DestructiveVerification"])), "Batch omitted structural-only verification");
            batch.DestructiveStyleCleanup = false;
            Check(batch.Items.All(item => item.StatusText == text["BatchPending"]), "Batch retained stale destructive results");
            CheckMissingNormalCounts(directory, text);
            Console.WriteLine("PASS: destructive GUI models, four-language warnings, option resets, structural-only verification and source preservation.");
        }
        finally { Directory.Delete(directory, true); }
    }

    private static void CheckMissingNormalCounts(string directory, UiText text)
    {
        foreach (var omitNames in new[] { false, true })
        {
            var input = Path.Combine(directory, omitNames ? "no-names.xlsx" : "no-normal.xlsx");
            ExcelToolChecks.CreateNamedStyleWorkbook(input);
            using (var archive = ZipFile.Open(input, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("xl/styles.xml")!;
                XDocument styles;
                using (var stream = entry.Open()) styles = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var names = styles.Root!.Element(ns + "cellStyles")!;
                if (omitNames) names.Remove();
                else
                {
                    names.Elements().Single(name => (string?)name.Attribute("builtinId") == "0").Remove();
                    names.SetAttributeValue("count", names.Elements().Count());
                }
                entry.Delete();
                using var output = archive.CreateEntry("xl/styles.xml").Open();
                styles.Save(output);
            }
            var model = new ExcelToolModel(text) { FilePath = input, DestructiveStyleCleanup = true };
            var analysis = new ExcelStyleAnalyzer().Analyze(input, options: model.OptimizationOptions);
            model.ApplyAnalysis(analysis);
            var expected = omitNames ? "0" : "2";
            var label = text["DestructiveNamedStyleRemovals"];
            Check(model.Summary!.Diagnostics.Single(row => row.Label == label).Value == expected,
                "Missing Normal analysis must count discarded names, not net table size change");
            var destination = Path.Combine(directory, omitNames ? "no-names-output" : "no-normal-output");
            var file = new ExcelBatchFile(input, Path.GetFileName(input), new FileInfo(input).Length);
            var outcome = new ExcelBatchOptimizer().Optimize([file], destination, overwrite: false, options: model.OptimizationOptions);
            var result = outcome.Results.Single();
            Check(result.Status == ExcelBatchItemStatus.Completed, "Missing Normal batch fixture failed: " + result.Error);
            var item = new ExcelBatchItem(file, text);
            item.Apply(result, text);
            Check(omitNames ? !item.DetailText.Contains(label) : item.DetailText.Contains(label + ": " + expected),
                "Missing Normal batch must count all discarded names");
            var optimized = new ExcelStyleOptimizer().Optimize(input, Path.Combine(directory, Path.GetFileNameWithoutExtension(input) + "-single.xlsx"),
                options: model.OptimizationOptions);
            model.ApplyOptimization(optimized);
            Check(model.Result!.Metrics.Single(row => row.Label == label).Value == expected,
                "Missing Normal output must count discarded names, not net table size change");
        }
    }

    private static void Check(bool value, string message)
    {
        if (!value) throw new Exception(message);
    }
}
