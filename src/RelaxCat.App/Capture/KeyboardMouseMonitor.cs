using System.ComponentModel;
using System.Runtime.InteropServices;
using RelaxCat.Core.Capture;

namespace RelaxCat.App.Capture;

// Construct/dispose on the Avalonia UI thread, whose Windows message loop keeps the LL hook alive.
public sealed class KeyboardMouseMonitor : IDisposable
{
    private readonly HookCallback callback;
    private readonly CaptureOptions options;
    private nint hook;
    private delegate nint HookCallback(int code, nint message, nint data);
    public KeyboardMouseMonitor(CaptureOptions options)
    {
        this.options = options; callback = OnKey;
        hook = SetWindowsHookEx(13, callback, GetModuleHandle(null), 0);
        if (hook == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
    }
    private nint OnKey(int code, nint message, nint data)
    {
        try
        {
            if (code >= 0 && (message == 0x100 || message == 0x104))
            {
                var key = Marshal.ReadInt32(data);
                var modifiers = ShortcutModifiers.None;
                if (Down(0x11)) modifiers |= ShortcutModifiers.Control;
                if (Down(0x12)) modifiers |= ShortcutModifiers.Alt;
                if (Down(0x10)) modifiers |= ShortcutModifiers.Shift;
                if (Down(0x5B) || Down(0x5C)) modifiers |= ShortcutModifiers.Windows;
                if (options.MonitoredKeys.Any(shortcut => shortcut.VirtualKey == key && shortcut.Modifiers == modifiers))
                    SetCursorPos(options.MouseX, options.MouseY);
            }
        }
        catch { /* Always forward the same key, even if pointer movement fails. */ }
        return CallNextHookEx(hook, code, message, data);
    }
    private static bool Down(int key) => (GetAsyncKeyState(key) & 0x8000) != 0;
    public void Dispose() { if (hook != 0) { UnhookWindowsHookEx(hook); hook = 0; } }
    [DllImport("user32.dll", SetLastError = true)] private static extern nint SetWindowsHookEx(int id, HookCallback callback, nint module, uint threadId);
    [DllImport("user32.dll")] private static extern nint CallNextHookEx(nint hook, int code, nint message, nint data);
    [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(nint hook);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern nint GetModuleHandle(string? name);
}
