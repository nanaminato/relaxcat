using System.Text.Json;
using System.Text;
using RelaxCat.App.Models;
using RelaxCat.App.Views;
using Avalonia.Controls;
using Avalonia.Threading;

namespace RelaxCat.App.SmokeTests;

internal static class LocalizationChecks
{
    internal static void Run()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "Localization");
        var chinese = Read(Path.Combine(directory, "zh-CN.json"));
        var text = new UiText();
        foreach (var language in text.AvailableLanguages)
        {
            var strings = Read(Path.Combine(directory, language.Culture + ".json"));
            if (!chinese.Keys.Order().SequenceEqual(strings.Keys.Order())) throw new Exception("Language pack key sets differ: " + language.Culture);
            text.CurrentCulture = language.Culture;
            foreach (var (key, value) in strings)
            {
                if (text[key] != value) throw new Exception("Resource did not load: " + language.Culture + "/" + key);
                CompositeFormat.Parse(value);
            }
        }
        text.CurrentCulture = UiText.DefaultCulture;
        var notifications = new List<string?>();
        text.PropertyChanged += (_, e) => notifications.Add(e.PropertyName);
        text.CurrentCulture = "ja-JP";
        foreach (var (key, value) in Read(Path.Combine(directory, "ja-JP.json")))
            if (text[key] != value) throw new Exception("Japanese resource did not load: " + key);
        if (!notifications.Contains("Item[]")) throw new Exception("Language switch did not notify indexer bindings");

        var temp = Path.GetFullPath("artifacts/localization-checks-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        try
        {
            WritePack(temp, "zh-CN", "简体中文", 0, "{\"Existing\":\"baseline\"}");
            WritePack(temp, "ja-JP", "日本語", 10, "{\"Blank\":\" \"}");
            var fallback = new UiText(temp) { CurrentCulture = "ja-JP" };
            if (fallback["Existing"] != "baseline" || fallback["Missing"] != "Missing") throw new Exception("Locale fallback failed");
            WritePack(temp, "ja-JP", "日本語", 10, "{\"Existing\":\"invalid {0\"}");
            if (new UiText(temp) { CurrentCulture = "ja-JP" }["Existing"] != "baseline") throw new Exception("Invalid format fallback failed");
            File.WriteAllText(Path.Combine(temp, "ja-JP.json"), "{bad json");
            if (new UiText(temp) { CurrentCulture = "ja-JP" }["Existing"] != "baseline") throw new Exception("Malformed optional pack fallback failed");
            File.Delete(Path.Combine(temp, "ja-JP.json"));
            if (new UiText(temp) { CurrentCulture = "ja-JP" }["Existing"] != "baseline") throw new Exception("Missing optional pack fallback failed");
            CheckDynamicLanguages(temp);
        }
        finally { Directory.Delete(temp, true); }
        Console.WriteLine($"PASS: external JSON packs ({chinese.Count} keys), language switch, placeholders and fallback.");
    }

    private static void CheckDynamicLanguages(string directory)
    {
        WritePack(directory, "zh-CN", "简体中文", 0, "{\"Existing\":\"baseline\"}");
        WritePack(directory, "ja-JP", "日本語", 10, "{\"Existing\":\"日本語\"}");
        WritePack(directory, "en-US", "English", 20, "{\"Existing\":\"English value\"}");
        WritePack(directory, "ko-KR", "한국어", 30, "{\"Existing\":\"한국어 값\"}");
        var catalog = new UiText(directory);
        if (catalog.AvailableLanguages.Count != 4) throw new Exception("New JSON languages were not discovered");
        var settingsPath = Path.Combine(directory, "preferences.tmp"); // not a language-pack JSON
        using (var workspace = new WorkspaceModel(new SettingsStore(settingsPath), catalog))
        {
            var view = new SettingsView { DataContext = workspace };
            var window = new Window { Content = view, Width = 900, Height = 700 };
            window.Show(); Dispatcher.UIThread.RunJobs();
            var selector = view.FindControl<ComboBox>("LanguageSelector")!;
            if (selector.Items.Count != 4) throw new Exception("Selector still has a fixed language list");
            selector.SelectedIndex = catalog.AvailableLanguages.ToList().FindIndex(l => l.Culture == "en-US");
            Dispatcher.UIThread.RunJobs();
            if (catalog.CurrentCulture != "en-US" || catalog["Existing"] != "English value") throw new Exception("English selection failed");
            selector.SelectedIndex = catalog.AvailableLanguages.ToList().FindIndex(l => l.Culture == "ko-KR");
            Dispatcher.UIThread.RunJobs();
            if (catalog.CurrentCulture != "ko-KR" || catalog["Existing"] != "한국어 값") throw new Exception("Korean selection failed");
            workspace.FlushSettings();
            window.Close(); Dispatcher.UIThread.RunJobs();
        }
        // Reordering/adding language files must not change a persisted language.
        WritePack(directory, "ko-KR", "한국어", -10, "{\"Existing\":\"한국어 값\"}");
        using (var restored = new WorkspaceModel(new SettingsStore(settingsPath), new UiText(directory)))
            if (restored.Text.CurrentCulture != "ko-KR" || restored.LanguageIndex != 0) throw new Exception("Language was persisted by list index");
        var saved = new SettingsStore.Snapshot(new SettingsModel(), new CaptureSettings(), 1, 0);
        File.WriteAllText(settingsPath, JsonSerializer.Serialize(saved));
        using (var legacy = new WorkspaceModel(new SettingsStore(settingsPath), new UiText(directory)))
            if (legacy.Text.CurrentCulture != "ja-JP") throw new Exception("Legacy two-language preferences were not migrated");
        File.Delete(Path.Combine(directory, "ko-KR.json"));
        catalog = new UiText(directory) { CurrentCulture = "ko-KR" };
        if (catalog.CurrentCulture != UiText.DefaultCulture) throw new Exception("Missing language did not fall back to baseline");
        Console.WriteLine("PASS: JSON-only English/Korean discovery, actual selector switching, culture persistence/reordering and legacy migration.");
    }

    private static void WritePack(string directory, string culture, string name, int order, string strings) =>
        File.WriteAllText(Path.Combine(directory, culture + ".json"), JsonSerializer.Serialize(new
        { culture, displayName = name, sortOrder = order, strings = JsonSerializer.Deserialize<Dictionary<string, string>>(strings) }));

    private static Dictionary<string, string> Read(string path) => JsonSerializer.Deserialize<Dictionary<string, string>>(JsonDocument.Parse(File.ReadAllText(path)).RootElement.GetProperty("strings").GetRawText())!;
}
