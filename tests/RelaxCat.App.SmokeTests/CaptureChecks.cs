using System.Diagnostics;
using System.Runtime.InteropServices;
using Avalonia.Controls;
using Avalonia.Interactivity;
using RelaxCat.App;
using RelaxCat.App.Capture;
using RelaxCat.App.Views;

internal static class CaptureChecks
{
    public static void RunFolder(MainWindow window, Action<Task> complete, Action<Func<bool>> until)
    {
        var folder = Path.GetFullPath("artifacts/监控 文件夹-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        File.Copy("artifacts/sample-1.png", Path.Combine(folder, "old.png"));
        var count = window.Workspace.Images.Count;
        window.Workspace.Capture.ProviderIndex = 1; window.Workspace.Capture.Folder = folder;
        window.FindControl<Button>("CaptureSettingsButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        var view = window.FindControl<SettingsView>("SettingsPage")!;
        view.FindControl<Button>("ApplyCaptureButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        until(() => !window.Workspace.CaptureBusy);
        Check(window.Workspace.CaptureRunning, "Folder provider did not start");
        window.Workspace.IsBusy = true;
        File.Copy("artifacts/sample-2.png", Path.Combine(folder, "queued.png"));
        var delay = Task.Delay(1200); complete(delay);
        Check(window.Workspace.Images.Count == count, "Capture was not queued behind busy workflow");
        window.Workspace.IsBusy = false;
        until(() => window.Workspace.Images.Count == count + 1);
        complete(Task.Delay(900));
        Check(window.Workspace.Images.Count == count + 1, "Duplicate capture or existing backlog was imported");
        view.FindControl<Button>("StopCaptureButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        until(() => !window.Workspace.CaptureBusy);
        Check(!window.Workspace.CaptureRunning, "Folder provider did not stop");
        File.Copy("artifacts/sample-3.png", Path.Combine(folder, "after-stop.png"));
        complete(Task.Delay(900)); Check(window.Workspace.Images.Count == count + 1, "Import continued after stopping");
        Console.WriteLine("PASS: real folder monitoring through XAML controls, baseline exclusion, busy-queue delivery, deduplication and stop.");
    }

    public static int RunGreenshot(MainWindow window, Action<Task> complete, Action<Func<bool>> until)
    {
        var folder = Path.GetFullPath("artifacts/Greenshot 截图-" + Guid.NewGuid().ToString("N"));
        window.CaptureStateRoot = Path.GetFullPath("artifacts/greenshot-state-" + Guid.NewGuid().ToString("N"));
        var capture = window.Workspace.Capture;
        capture.ProviderIndex = 2; capture.Folder = folder;
        capture.RegionKey = "F22"; capture.WindowKey = "F23"; capture.FullscreenKey = "F24";
        try
        {
            var cancelledStart = window.ApplyCaptureAsync();
            complete(window.StopCaptureAsync()); complete(cancelledStart);
            Check(!window.Workspace.CaptureRunning && !window.GreenshotProcessId.HasValue, "Cancelled startup left a provider running");
            complete(window.ApplyCaptureAsync());
            Check(window.Workspace.CaptureRunning && window.GreenshotProcessId.HasValue, window.Workspace.Message);
            using var first = Process.GetProcessById(window.GreenshotProcessId!.Value);
            using var competing = new AsyncProvider(new GreenshotProvider());
            try
            {
                complete(competing.Provider.StartAsync(capture.CreateOptions(window.Workspace.Text), window.CaptureStateRoot, default));
                throw new Exception("An existing Greenshot was taken over");
            }
            catch (InvalidOperationException) { Check(!first.HasExited, "Existing instance was stopped"); }
            var before = window.Workspace.Images.Count;
            keybd_event(0x87, 0, 0, 0); keybd_event(0x87, 0, 2, 0); // F24 down/up, not the user's PrintScreen binding.
            until(() => window.Workspace.Images.Count == before + 1 || !window.Workspace.CaptureRunning);
            Check(window.Workspace.Images.Count == before + 1, window.Workspace.Message);
            Check(window.Workspace.Images[^1].Image.Path.StartsWith(folder, StringComparison.OrdinalIgnoreCase), "Greenshot did not save into monitored folder");
            capture.ProviderIndex = 1; complete(window.ApplyCaptureAsync());
            Check(first.HasExited && window.Workspace.CaptureRunning, "Provider switch did not stop only the owned Greenshot");
            capture.ProviderIndex = 2; complete(window.ApplyCaptureAsync());
            Check(window.GreenshotProcessId.HasValue, window.Workspace.Message);
            using (var crash = Process.GetProcessById(window.GreenshotProcessId!.Value)) crash.Kill();
            until(() => !window.Workspace.CaptureRunning && window.Workspace.Message == window.Workspace.Text["CaptureGreenshotExited"]);
            complete(window.ApplyCaptureAsync());
            Check(window.GreenshotProcessId.HasValue, window.Workspace.Message);
            Console.WriteLine("PASS: verified bundled Greenshot startup, own INI, F24 screenshot using Greenshot hotkeys, file-to-gallery pipeline, external-instance refusal, provider switch, unexpected-exit recovery.");
            return window.GreenshotProcessId!.Value; // Keep running to verify main-window shutdown cleanup.
        }
        catch { complete(window.StopCaptureAsync()); throw; }
    }
    private sealed class AsyncProvider(GreenshotProvider provider) : IDisposable
    {
        public GreenshotProvider Provider { get; } = provider;
        public void Dispose() => Provider.DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }

    public static void PumpNativeMessages()
    {
        while (PeekMessage(out var message, 0, 0, 0, 1)) { TranslateMessage(ref message); DispatchMessage(ref message); }
    }
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] private struct Message
    { public nint Window; public uint Id; public nuint WParam; public nint LParam; public uint Time; public Point Position; public uint Private; }
    [DllImport("user32.dll")] private static extern bool PeekMessage(out Message message, nint window, uint min, uint max, uint remove);
    [DllImport("user32.dll")] private static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll")] private static extern nint DispatchMessage(ref Message message);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, nuint extraInfo);
}
