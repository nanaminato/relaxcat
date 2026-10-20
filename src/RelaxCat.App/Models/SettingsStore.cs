using System.Text.Json;
using RelaxCat.App.Capture;

namespace RelaxCat.App.Models;

public sealed class SettingsStore
{
    private readonly string path;
    public bool RecoveredFromBackup { get; private set; }
    public SettingsStore(string? path = null) => this.path = path ?? Path.Combine(GreenshotBundle.DefaultStateRoot, "settings.json");
    public sealed record Snapshot(SettingsModel Settings, CaptureSettings Capture, int LanguageIndex, int ThemeIndex, int SchemaVersion = 1);
    private static Snapshot Read(string file)
    {
        var snapshot = JsonSerializer.Deserialize<Snapshot>(File.ReadAllText(file));
        if (snapshot is null || snapshot.Settings is null || snapshot.Capture is null || snapshot.SchemaVersion is < 0 or > 1)
            throw new JsonException("Unsupported or incomplete settings file.");
        foreach (var model in new object[] { snapshot.Settings, snapshot.Capture })
            if (model.GetType().GetProperties().Any(p => p.PropertyType == typeof(string) && p.GetValue(model) is null))
                throw new JsonException("A settings value is missing.");
        return snapshot;
    }
    public Snapshot? Load()
    {
        RecoveredFromBackup = false;
        if (!File.Exists(path) && !File.Exists(path + ".bak")) return null;
        try { return Read(path); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException)
        {
            var restored = Read(path + ".bak");
            RecoveredFromBackup = true;
            return restored;
        }
    }
    public void Save(WorkspaceModel model)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, JsonSerializer.Serialize(new Snapshot(model.Settings, model.Capture, model.LanguageIndex, model.ThemeIndex), new JsonSerializerOptions { WriteIndented = true }));
        // Never replace a healthy backup with a corrupt primary file.
        if (File.Exists(path))
        {
            try { Read(path); File.Copy(path, path + ".bak", true); }
            catch (JsonException) { }
        }
        File.Move(temporary, path, true);
    }
}
