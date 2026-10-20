namespace RelaxCat.Core.Capture;

[Flags]
public enum ShortcutModifiers { None = 0, Control = 1, Alt = 2, Shift = 4, Windows = 8 }

public readonly record struct CaptureShortcut(int VirtualKey, ShortcutModifiers Modifiers, string KeyName)
{
    public static CaptureShortcut Parse(string text)
    {
        var parts = text.Split('+', StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || parts.Any(string.IsNullOrWhiteSpace)) throw new FormatException("A shortcut is required.");
        var modifiers = ShortcutModifiers.None;
        foreach (var modifier in parts[..^1])
        {
            var value = modifier.ToUpperInvariant() switch
            {
                "CTRL" or "CONTROL" => ShortcutModifiers.Control,
                "ALT" => ShortcutModifiers.Alt,
                "SHIFT" => ShortcutModifiers.Shift,
                "WIN" or "WINDOWS" => ShortcutModifiers.Windows,
                _ => throw new FormatException($"Unknown modifier: {modifier}")
            };
            if ((modifiers & value) != 0) throw new FormatException($"Duplicate modifier: {modifier}");
            modifiers |= value;
        }
        var key = parts[^1].ToUpperInvariant();
        var virtualKey = key switch
        {
            "PRINTSCREEN" or "PRTSC" or "SNAPSHOT" => 0x2C,
            _ when key.Length == 1 && char.IsAsciiLetterOrDigit(key[0]) => key[0],
            _ when key.StartsWith('F') && int.TryParse(key[1..], out var f) && f is >= 1 and <= 24 => 0x70 + f - 1,
            _ => throw new FormatException($"Unsupported shortcut key: {key}")
        };
        return new CaptureShortcut(virtualKey, modifiers, virtualKey == 0x2C ? "PrintScreen" : key);
    }

    public string ToGreenshotString() => string.Join(" + ", new[]
    {
        Modifiers.HasFlag(ShortcutModifiers.Control) ? "Ctrl" : null,
        Modifiers.HasFlag(ShortcutModifiers.Alt) ? "Alt" : null,
        Modifiers.HasFlag(ShortcutModifiers.Shift) ? "Shift" : null,
        Modifiers.HasFlag(ShortcutModifiers.Windows) ? "Win" : null,
        KeyName
    }.Where(part => part is not null));
}
