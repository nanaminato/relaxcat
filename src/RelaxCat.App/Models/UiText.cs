using System.Diagnostics;
using System.Text.Json;
using System.Text;

namespace RelaxCat.App.Models;

/// <summary>Stable resource keys backed by JSON files shipped beside the application.</summary>
public sealed class UiText : ObservableModel
{
    private static readonly Lazy<LanguageCatalog> DefaultCatalog = new(() =>
        LanguageCatalog.Load(Path.Combine(AppContext.BaseDirectory, "Localization")));
    private readonly LanguageCatalog catalog;
    private bool japanese;

    public UiText() : this(DefaultCatalog.Value) { }
    internal UiText(string directory) : this(LanguageCatalog.Load(directory)) { }
    private UiText(LanguageCatalog catalog) => this.catalog = catalog;

    public bool Japanese
    {
        get => japanese;
        set { if (Set(ref japanese, value)) { Notify("Item"); Notify("Item[]"); } }
    }

    // Chinese is the baseline; an absent/blank Japanese entry falls back to it, then to the key.
    public string this[string key] => catalog.Get(key, Japanese);

    private sealed record LanguageCatalog(IReadOnlyDictionary<string, string> Chinese, IReadOnlyDictionary<string, string> Japanese)
    {
        internal static LanguageCatalog Load(string directory)
        {
            var chinese = Read(Path.Combine(directory, "zh-CN.json"));
            IReadOnlyDictionary<string, string> japanese;
            try { japanese = Read(Path.Combine(directory, "ja-JP.json")); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or InvalidDataException)
            {
                Trace.TraceWarning("Cannot load ja-JP language pack; using zh-CN. {0}", error.Message);
                japanese = new Dictionary<string, string>();
            }
            // External packs can lag behind the baseline. Reject invalid composite formats
            // per key so a translator cannot break a later string.Format call.
            japanese = japanese.Where(entry => chinese.TryGetValue(entry.Key, out var baseline) &&
                CompatibleFormat(baseline, entry.Value)).ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.Ordinal);
            return new LanguageCatalog(chinese, japanese);
        }

        internal string Get(string key, bool japanese)
        {
            if (japanese && Japanese.TryGetValue(key, out var translated) && !string.IsNullOrWhiteSpace(translated)) return translated;
            return Chinese.TryGetValue(key, out var baseline) && !string.IsNullOrWhiteSpace(baseline) ? baseline : key;
        }

        private static bool CompatibleFormat(string baseline, string translation)
        {
            try { return CompositeFormat.Parse(baseline).MinimumArgumentCount == CompositeFormat.Parse(translation).MinimumArgumentCount; }
            catch (FormatException) { return false; }
        }

        private static IReadOnlyDictionary<string, string> Read(string path)
        {
            using var stream = File.OpenRead(path);
            using var document = JsonDocument.Parse(stream);
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException($"Language pack '{path}' must contain an object of string entries.");
            var entries = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var entry in document.RootElement.EnumerateObject())
                if (entry.Value.ValueKind != JsonValueKind.String || !entries.TryAdd(entry.Name, entry.Value.GetString()!))
                    throw new InvalidDataException($"Invalid or duplicate resource key '{entry.Name}' in '{path}'.");
            return entries;
        }
    }
}
