using RelaxCat.Core.Capture;

namespace RelaxCat.App.Capture;

public enum CaptureProviderKind { Off, Folder, Greenshot }
public sealed record CaptureOptions(CaptureProviderKind Provider, string Folder,
    CaptureShortcut RegionKey, CaptureShortcut WindowKey, CaptureShortcut FullscreenKey,
    string Language);
