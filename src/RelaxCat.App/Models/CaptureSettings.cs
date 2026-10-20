using System.Globalization;
using RelaxCat.App.Capture;
using RelaxCat.Core.Capture;

namespace RelaxCat.App.Models;

public sealed class CaptureSettings : ObservableModel
{
    private int providerIndex;
    private string folder = Path.Combine(GreenshotBundle.DefaultStateRoot, "Captures");
    private string regionKey = "PrintScreen", windowKey = "Alt+PrintScreen", fullscreenKey = "Ctrl+PrintScreen";
    private string monitoredKeys = "PrintScreen, Alt+PrintScreen, Ctrl+PrintScreen", mouseX = "0", mouseY = "0";
    private bool moveMouse;
    public int ProviderIndex { get => providerIndex; set { if (value is >= 0 and <= 2 && Set(ref providerIndex, value)) Notify(nameof(IsGreenshot)); } }
    public bool IsGreenshot => ProviderIndex == 2;
    public string Folder { get => folder; set => Set(ref folder, value); }
    public string RegionKey { get => regionKey; set => Set(ref regionKey, value); }
    public string WindowKey { get => windowKey; set => Set(ref windowKey, value); }
    public string FullscreenKey { get => fullscreenKey; set => Set(ref fullscreenKey, value); }
    public string MonitoredKeys { get => monitoredKeys; set => Set(ref monitoredKeys, value); }
    public string MouseX { get => mouseX; set => Set(ref mouseX, value); }
    public string MouseY { get => mouseY; set => Set(ref mouseY, value); }
    public bool MoveMouse { get => moveMouse; set => Set(ref moveMouse, value); }
    public CaptureOptions CreateOptions(UiText text)
    {
        var provider = (CaptureProviderKind)ProviderIndex;
        var basic = new CaptureOptions(provider, "", CaptureShortcut.Parse("PrintScreen"), CaptureShortcut.Parse("Alt+PrintScreen"),
            CaptureShortcut.Parse("Ctrl+PrintScreen"), false, 0, 0, Array.Empty<CaptureShortcut>(), text.Japanese ? "ja-JP" : "zh-CN");
        if (provider == CaptureProviderKind.Off) return basic;
        if (string.IsNullOrWhiteSpace(Folder) || Folder.IndexOfAny(new[] { '\r', '\n' }) >= 0) throw new ArgumentException(text["CaptureFolderRequired"]);
        var result = basic with { Folder = Path.GetFullPath(Folder) };
        if (provider == CaptureProviderKind.Folder) return result;
        var region = CaptureShortcut.Parse(RegionKey); var window = CaptureShortcut.Parse(WindowKey); var fullscreen = CaptureShortcut.Parse(FullscreenKey);
        var configured = new[] { region, window, fullscreen };
        if (configured.Select(key => (key.VirtualKey, key.Modifiers)).Distinct().Count() != 3) throw new ArgumentException(text["CaptureKeysDistinct"]);
        foreach (var key in configured)
            if ((key.VirtualKey is >= 0x30 and <= 0x39 or >= 0x41 and <= 0x5A) && (key.Modifiers == ShortcutModifiers.None || key.Modifiers == ShortcutModifiers.Shift))
                throw new ArgumentException(text["CaptureKeyModifier"]);
        var monitored = MoveMouse ? MonitoredKeys.Split(new[] { ',', ';', '\n' }, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).Select(CaptureShortcut.Parse).ToArray()
            : Array.Empty<CaptureShortcut>();
        if (MoveMouse && monitored.Length == 0) throw new ArgumentException(text["CaptureMonitorKeysRequired"]);
        var x = 0; var y = 0;
        if (MoveMouse && (!int.TryParse(MouseX, NumberStyles.Integer, CultureInfo.InvariantCulture, out x) || !int.TryParse(MouseY, NumberStyles.Integer, CultureInfo.InvariantCulture, out y)))
            throw new ArgumentException(text["CaptureCoordinatesInvalid"]);
        return result with { RegionKey = region, WindowKey = window, FullscreenKey = fullscreen, MoveMouse = MoveMouse, MouseX = x, MouseY = y, MonitoredKeys = monitored };
    }
}
