using System.Text.Json;
using System.Text;
using RelaxCat.App.Models;

namespace RelaxCat.App.SmokeTests;

internal static class LocalizationChecks
{
    internal static void Run()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "Localization");
        var chinese = Read(Path.Combine(directory, "zh-CN.json"));
        var japanese = Read(Path.Combine(directory, "ja-JP.json"));
        if (!chinese.Keys.Order().SequenceEqual(japanese.Keys.Order())) throw new Exception("Language pack key sets differ");
        var text = new UiText();
        foreach (var (key, value) in chinese)
        {
            if (text[key] != value) throw new Exception("Chinese resource did not load: " + key);
            if (CompositeFormat.Parse(value).MinimumArgumentCount != CompositeFormat.Parse(japanese[key]).MinimumArgumentCount)
                throw new Exception("Language format parameters differ: " + key);
        }
        var notifications = new List<string?>();
        text.PropertyChanged += (_, e) => notifications.Add(e.PropertyName);
        text.Japanese = true;
        foreach (var (key, value) in japanese)
            if (text[key] != value) throw new Exception("Japanese resource did not load: " + key);
        if (!notifications.Contains("Item[]")) throw new Exception("Language switch did not notify indexer bindings");

        var temp = Path.GetFullPath("artifacts/localization-checks-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        try
        {
            File.WriteAllText(Path.Combine(temp, "zh-CN.json"), "{\"Existing\":\"baseline\"}");
            File.WriteAllText(Path.Combine(temp, "ja-JP.json"), "{\"Blank\":\" \"}");
            var fallback = new UiText(temp) { Japanese = true };
            if (fallback["Existing"] != "baseline" || fallback["Missing"] != "Missing") throw new Exception("Locale fallback failed");
            File.WriteAllText(Path.Combine(temp, "ja-JP.json"), "{\"Existing\":\"invalid {0\"}");
            if (new UiText(temp) { Japanese = true }["Existing"] != "baseline") throw new Exception("Invalid format fallback failed");
            File.WriteAllText(Path.Combine(temp, "ja-JP.json"), "{bad json");
            if (new UiText(temp) { Japanese = true }["Existing"] != "baseline") throw new Exception("Malformed optional pack fallback failed");
            File.Delete(Path.Combine(temp, "ja-JP.json"));
            if (new UiText(temp) { Japanese = true }["Existing"] != "baseline") throw new Exception("Missing optional pack fallback failed");
        }
        finally { Directory.Delete(temp, true); }
        Console.WriteLine($"PASS: external JSON packs ({chinese.Count} keys), language switch, placeholders and fallback.");
    }

    private static Dictionary<string, string> Read(string path) => JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path))!;
}
