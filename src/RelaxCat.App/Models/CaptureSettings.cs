using RelaxCat.App.Capture;
using RelaxCat.Core.Capture;

namespace RelaxCat.App.Models;

public sealed class CaptureSettings : ObservableModel
{
    private int providerIndex;
    private string folder = Path.Combine(GreenshotBundle.DefaultStateRoot, "Captures");
    private string regionKey = "PrintScreen", windowKey = "Alt+PrintScreen", fullscreenKey = "Ctrl+PrintScreen";
    public int ProviderIndex { get => providerIndex; set { if (value is >= 0 and <= 2 && Set(ref providerIndex, value)) Notify(nameof(IsGreenshot)); } }
    public bool IsGreenshot => ProviderIndex == 2;
    public string Folder { get => folder; set => Set(ref folder, value); }
    public string RegionKey { get => regionKey; set => Set(ref regionKey, value); }
    public string WindowKey { get => windowKey; set => Set(ref windowKey, value); }
    public string FullscreenKey { get => fullscreenKey; set => Set(ref fullscreenKey, value); }
    public CaptureOptions CreateOptions(UiText text)
    {
        var provider = (CaptureProviderKind)ProviderIndex;
        var basic = new CaptureOptions(provider, "", CaptureShortcut.Parse("PrintScreen"), CaptureShortcut.Parse("Alt+PrintScreen"),
            CaptureShortcut.Parse("Ctrl+PrintScreen"), text.Japanese ? "ja-JP" : "zh-CN");
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
        return result with { RegionKey = region, WindowKey = window, FullscreenKey = fullscreen };
    }
}
