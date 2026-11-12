using System.Text.Json;
using RelaxCat.App.Models;
using RelaxCat.Core.Excel;

namespace RelaxCat.App.SmokeTests;

internal static class OutputNamingChecks
{
    internal static void Run()
    {
        var settings = new SettingsModel();
        Check(!settings.AppendOptimizationSuffix, "New installations must keep original names");
        settings.AppendOptimizationSuffix = true;
        Check(JsonSerializer.Deserialize<SettingsModel>(JsonSerializer.Serialize(settings))!.AppendOptimizationSuffix,
            "Naming preference must survive settings serialization");
        Check(!JsonSerializer.Deserialize<SettingsModel>("{}")!.AppendOptimizationSuffix,
            "Old settings must default to original names");
        var text = new UiText();
        foreach (var language in text.AvailableLanguages)
        {
            text.CurrentCulture = language.Culture;
            Check(text["AppendOptimizationSuffix"] != "AppendOptimizationSuffix" &&
                  text["OptimizationNamingHint"] != "OptimizationNamingHint", "Missing localized naming controls");
        }
        var batch = new ExcelBatchModel(text) { InputFolder = Path.GetTempPath(), OutputFolder = "first-output" };
        batch.ApplyDiscovery(new ExcelBatchDiscovery([new ExcelBatchFile("input.xlsx", "input.xlsx", 1)], []));
        Check(batch.HasItems && batch.CanRun, "Test scan failed");
        batch.OutputFolder = "second-output";
        Check(!batch.HasItems && !batch.CanRun && batch.Files.Count == 0, "Output directory change reused stale scan");
        Console.WriteLine("PASS: output naming defaults, persistence, localization and stale batch scan invalidation.");
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
