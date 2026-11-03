using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace RelaxCat.App.Models;

public sealed record LanguageOption(string Culture, string DisplayName, int SortOrder);

/// <summary>Discovers external language packs; stable keys and culture IDs never depend on list order.</summary>
public sealed class UiText : ObservableModel
{
    public const string DefaultCulture = "zh-CN";
    private static readonly Lazy<LanguageCatalog> DefaultCatalog = new(() =>
        LanguageCatalog.Load(Path.Combine(AppContext.BaseDirectory, "Localization")));
    private readonly LanguageCatalog catalog;
    private string currentCulture = DefaultCulture;

    public UiText() : this(DefaultCatalog.Value) { }
    internal UiText(string directory) : this(LanguageCatalog.Load(directory)) { }
    private UiText(LanguageCatalog catalog) => this.catalog = catalog;
    public IReadOnlyList<LanguageOption> AvailableLanguages => catalog.Options;
    public string CurrentCulture
    {
        get => currentCulture;
        set
        {
            var resolved = catalog.Packs.Keys.FirstOrDefault(c => string.Equals(c, value, StringComparison.OrdinalIgnoreCase)) ?? DefaultCulture;
            if (Set(ref currentCulture, resolved)) { Notify("Item"); Notify("Item[]"); }
        }
    }
    public string this[string key] => catalog.Get(key, CurrentCulture);

    private sealed record LanguagePack(LanguageOption Option, IReadOnlyDictionary<string, string> Strings);
    private sealed class LanguageCatalog
    {
        internal required IReadOnlyDictionary<string, LanguagePack> Packs { get; init; }
        internal required IReadOnlyList<LanguageOption> Options { get; init; }
        internal static LanguageCatalog Load(string directory)
        {
            var baseline = Read(Path.Combine(directory, DefaultCulture + ".json"));
            var packs = new Dictionary<string, LanguagePack>(StringComparer.OrdinalIgnoreCase) { [DefaultCulture] = baseline };
            foreach (var path in Directory.EnumerateFiles(directory, "*.json").Order(StringComparer.Ordinal))
            {
                if (string.Equals(Path.GetFileNameWithoutExtension(path), DefaultCulture, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    var pack = Read(path);
                    var strings = pack.Strings.Where(entry => baseline.Strings.TryGetValue(entry.Key, out var source) &&
                        !string.IsNullOrWhiteSpace(entry.Value) && CompatibleFormat(source, entry.Value))
                        .ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.Ordinal);
                    if (!packs.TryAdd(pack.Option.Culture, pack with { Strings = strings }))
                        throw new InvalidDataException("Duplicate language culture: " + pack.Option.Culture);
                }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or InvalidDataException or CultureNotFoundException)
                { Trace.TraceWarning("Cannot load language pack '{0}'; skipping it. {1}", Path.GetFileName(path), error.Message); }
            }
            return new LanguageCatalog
            {
                Packs = packs,
                Options = packs.Values.Select(pack => pack.Option).OrderBy(option => option.SortOrder)
                    .ThenBy(option => option.Culture, StringComparer.Ordinal).ToArray()
            };
        }
        internal string Get(string key, string culture)
        {
            if (Packs[culture].Strings.TryGetValue(key, out var translated) && !string.IsNullOrWhiteSpace(translated)) return translated;
            return Packs[DefaultCulture].Strings.TryGetValue(key, out var fallback) ? fallback : key;
        }
        private static bool CompatibleFormat(string source, string translation)
        {
            try
            {
                CompositeFormat.Parse(source); CompositeFormat.Parse(translation);
                static HashSet<string> Parameters(string value) => Regex.Matches(value, @"(?<!\{)\{(\d+)(?:[,}:])")
                    .Select(match => match.Groups[1].Value).ToHashSet(StringComparer.Ordinal);
                return Parameters(source).SetEquals(Parameters(translation));
            }
            catch (FormatException) { return false; }
        }
        private static LanguagePack Read(string path)
        {
            using var stream = File.OpenRead(path);
            using var document = JsonDocument.Parse(stream);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) throw new InvalidDataException("Language pack must be an object.");
            if (root.EnumerateObject().Select(p => p.Name).Distinct(StringComparer.Ordinal).Count() != root.EnumerateObject().Count())
                throw new InvalidDataException("Duplicate language metadata.");
            if (!root.TryGetProperty("culture", out var cultureElement) || cultureElement.ValueKind != JsonValueKind.String ||
                !root.TryGetProperty("displayName", out var displayElement) || displayElement.ValueKind != JsonValueKind.String ||
                !root.TryGetProperty("strings", out var strings) || strings.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException("Language pack needs culture, displayName and strings.");
            var culture = CultureInfo.GetCultureInfo(cultureElement.GetString()!).Name;
            var displayName = displayElement.GetString()!;
            if (culture.Length == 0 || string.IsNullOrWhiteSpace(displayName) ||
                !string.Equals(culture, Path.GetFileNameWithoutExtension(path), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Language culture must match the filename and have a display name.");
            var order = 100;
            if (root.TryGetProperty("sortOrder", out var sort) && (sort.ValueKind != JsonValueKind.Number || !sort.TryGetInt32(out order)))
                throw new InvalidDataException("sortOrder must be an integer.");
            var entries = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var entry in strings.EnumerateObject())
                if (entry.Value.ValueKind != JsonValueKind.String || !entries.TryAdd(entry.Name, entry.Value.GetString()!))
                    throw new InvalidDataException("Invalid or duplicate resource key: " + entry.Name);
            return new LanguagePack(new LanguageOption(culture, displayName, order), entries);
        }
    }
}
