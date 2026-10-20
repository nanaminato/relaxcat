using System.Text.Json;
using RelaxCat.App.Capture;

namespace RelaxCat.App.Models;

public sealed class SettingsStore
{
    private readonly string path;
    public SettingsStore(string? path = null) => this.path = path ?? Path.Combine(GreenshotBundle.DefaultStateRoot, "settings.json");
    public sealed record Snapshot(SettingsModel Settings, CaptureSettings Capture, int LanguageIndex, int ThemeIndex);
    public Snapshot? Load() => File.Exists(path) ? JsonSerializer.Deserialize<Snapshot>(File.ReadAllText(path)) : null;
    public void Save(WorkspaceModel model)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, JsonSerializer.Serialize(new Snapshot(model.Settings, model.Capture, model.LanguageIndex, model.ThemeIndex), new JsonSerializerOptions { WriteIndented = true }));
        File.Move(temporary, path, true);
    }
}
