using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using RelaxCat.App.Capture;
using RelaxCat.Core.Capture;

namespace RelaxCat.App;

public sealed partial class MainWindow
{
    private CaptureFolderMonitor? folderMonitor;
    private GreenshotProvider? greenshot;
    private CancellationTokenSource? captureStartup;
    private int captureGeneration;
    private bool closeApproved;
    private bool closingCapture;
    internal string CaptureStateRoot { get; set; } = GreenshotBundle.DefaultStateRoot;
    internal int? GreenshotProcessId => greenshot?.OwnedProcessId;

    private void CaptureSettingsClicked(object? sender, RoutedEventArgs e)
    { Workspace.SettingsOpen = true; SettingsPage.OpenCapturePage(); }
    private async void ApplyCaptureRequested(object? sender, EventArgs e) => await ApplyCaptureAsync();
    private async void StopCaptureRequested(object? sender, EventArgs e) => await StopCaptureAsync();

    internal async Task ApplyCaptureAsync()
    {
        if (!Workspace.CanApplyCapture || closed) return;
        CaptureOptions options;
        try
        {
            if (Workspace.Capture.IsGreenshot && GreenshotProvider.ReadOutputFolder(CaptureStateRoot) is { } folder)
                Workspace.Capture.Folder = folder;
            options = Workspace.Capture.CreateOptions(Workspace.Text);
        }
        catch (Exception error) { Workspace.Message = Workspace.Text["CaptureError"] + error.Message; return; }
        Workspace.CaptureBusy = true; captureStartup = new CancellationTokenSource();
        try
        {
            await StopCaptureCoreAsync();
            if (options.Provider == CaptureProviderKind.Off) { Workspace.Message = Workspace.CaptureStatus; return; }
            Workspace.CaptureStatusKey = "CaptureStarting";
            var generation = ++captureGeneration;
            folderMonitor = new CaptureFolderMonitor(options.Folder, ImportCapturedAsync,
                error => Dispatcher.UIThread.Post(() => { if (!closed && generation == captureGeneration) Workspace.Message = Workspace.Text["CaptureError"] + error.Message; }));
            if (options.Provider == CaptureProviderKind.Greenshot)
            {
                greenshot = new GreenshotProvider();
                await greenshot.StartAsync(options, CaptureStateRoot, captureStartup.Token);
                greenshot.UnexpectedExit += () => Dispatcher.UIThread.Post(async () =>
                {
                    if (closed || generation != captureGeneration) return;
                    await StopCaptureAsync(); Workspace.Message = Workspace.Text["CaptureGreenshotExited"];
                });
            }
            Workspace.CaptureRunning = true;
            Workspace.CaptureStatusKey = options.Provider == CaptureProviderKind.Greenshot ? "CaptureGreenshotRunning" : "CaptureFolderRunning";
            Workspace.Message = Workspace.CaptureStatus + " · " + options.Folder;
        }
        catch (OperationCanceledException) { await StopCaptureCoreAsync(); Workspace.Message = Workspace.CaptureStatus; }
        catch (Exception error)
        {
            await StopCaptureCoreAsync();
            Workspace.Message = Workspace.Text["CaptureError"] + (error is GreenshotAlreadyRunningException ? Workspace.Text["CaptureOtherInstance"] : error.Message);
        }
        finally { captureStartup.Dispose(); captureStartup = null; Workspace.CaptureBusy = false; }
    }

    internal async Task StopCaptureAsync()
    {
        if (Workspace.CaptureBusy) { captureStartup?.Cancel(); return; }
        Workspace.CaptureBusy = true;
        try { await StopCaptureCoreAsync(); Workspace.Message = Workspace.CaptureStatus; }
        catch (Exception error) { Workspace.Message = Workspace.Text["CaptureError"] + error.Message; }
        finally { Workspace.CaptureBusy = false; }
    }
    private async Task StopCaptureCoreAsync()
    {
        captureGeneration++;
        var monitor = folderMonitor; folderMonitor = null;
        try { if (monitor is not null) await monitor.DisposeAsync(); }
        finally
        {
            var provider = greenshot; greenshot = null;
            try { if (provider is not null) await provider.DisposeAsync(); }
            finally { Workspace.CaptureRunning = false; Workspace.CaptureStatusKey = "CaptureStopped"; }
        }
    }

    private async Task<bool> ImportCapturedAsync(string path, CancellationToken token) => await Dispatcher.UIThread.InvokeAsync(async () =>
    {
        // Serialize with manual import/export instead of dropping captures while the UI is busy.
        while (Workspace.IsBusy && !closed) await Task.Delay(100, token);
        token.ThrowIfCancellationRequested();
        if (closed) return false;
        var before = Workspace.Images.Count;
        await ImportPathsAsync(new[] { path }, token);
        return Workspace.Images.Count > before;
    });

    private async void WindowClosing(object? sender, WindowClosingEventArgs e)
    {
        if (closeApproved) return;
        if (closingCapture) { e.Cancel = true; return; }
        if (operation is null && !Workspace.CaptureBusy && folderMonitor is null && greenshot is null) return;
        e.Cancel = true; closingCapture = true; IsEnabled = false;
        operation?.Cancel(); captureStartup?.Cancel();
        while (Workspace.CaptureBusy) await Task.Delay(50);
        await StopCaptureAsync();
        while (operation is not null) await Task.Delay(50);
        closeApproved = true; Close();
    }
}
