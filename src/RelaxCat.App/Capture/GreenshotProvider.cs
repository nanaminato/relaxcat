using System.Diagnostics;
using System.Text;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;

namespace RelaxCat.App.Capture;

public sealed class GreenshotAlreadyRunningException : InvalidOperationException
{
    public GreenshotAlreadyRunningException() : base("Another Greenshot is already running. Close it first or use folder-only monitoring.") { }
}

public sealed class GreenshotProvider : IAsyncDisposable
{
    private Process? owned;
    private OwnedProcessJob? job;
    private bool stopping;
    public event Action? UnexpectedExit;
    public int? OwnedProcessId => owned is { HasExited: false } process ? process.Id : null;
    public string? ExecutablePath { get; private set; }

    public async Task StartAsync(CaptureOptions options, string stateRoot, CancellationToken token)
    {
        var existing = Process.GetProcessesByName("Greenshot");
        try
        {
            if (existing.Any(process => !process.HasExited)) throw new GreenshotAlreadyRunningException();
        }
        finally { foreach (var process in existing) process.Dispose(); }
        ExecutablePath = await GreenshotBundle.EnsureAsync(stateRoot, token);
        var profile = Path.Combine(Path.GetFullPath(stateRoot), "GreenshotProfile", Environment.ProcessId.ToString());
        Directory.CreateDirectory(profile); Directory.CreateDirectory(options.Folder);
        await File.WriteAllTextAsync(Path.Combine(profile, "greenshot.ini"), BuildConfiguration(options), new UTF8Encoding(false), token);
        var start = new ProcessStartInfo(ExecutablePath)
        {
            UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(ExecutablePath)!
        };
        start.ArgumentList.Add("--ini-directory"); start.ArgumentList.Add(profile);
        start.ArgumentList.Add("--autostart"); // A racing external instance must not open an instance-selection dialog.
        owned = Process.Start(start) ?? throw new IOException("Greenshot failed to start.");
        stopping = false;
        try
        {
            job = new OwnedProcessJob(owned);
            // Connecting without sending commands never triggers the CLI's auto-start behavior.
            var deadline = DateTime.UtcNow.AddSeconds(20);
            using var identity = WindowsIdentity.GetCurrent();
            var pipeName = "Greenshot_" + (identity.User?.Value ?? throw new IOException("Windows user SID unavailable."));
            var ready = false;
            while (DateTime.UtcNow < deadline)
            {
                token.ThrowIfCancellationRequested();
                if (owned.HasExited) throw new IOException($"Greenshot exited during startup ({owned.ExitCode}).");
                using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                try
                {
                    await pipe.ConnectAsync(250, token);
                    if (!GetNamedPipeServerProcessId(pipe.SafePipeHandle, out var serverPid) || serverPid != owned.Id)
                        throw new IOException("Greenshot IPC endpoint belongs to another process.");
                    ready = true; break;
                }
                catch (TimeoutException) { }
                await Task.Delay(150, token);
            }
            if (!ready) throw new TimeoutException("Greenshot did not become ready.");
            owned.EnableRaisingEvents = true;
            owned.Exited += (_, _) => { if (!stopping) UnexpectedExit?.Invoke(); };
            if (owned.HasExited) throw new IOException("Greenshot exited during startup.");
        }
        catch { await DisposeAsync(); throw; }
    }

    internal static string BuildConfiguration(CaptureOptions options) => $$"""
        [Core]
        Language={{options.Language}}
        WelcomeShown=True
        RegionHotkey={{options.RegionKey.ToGreenshotString()}}
        WindowHotkey={{options.WindowKey.ToGreenshotString()}}
        FullscreenHotkey={{options.FullscreenKey.ToGreenshotString()}}
        LastregionHotkey=None
        ClipboardHotkey=None
        Destinations=FileNoDialog
        OutputFilePath={{options.Folder}}
        OutputFileFilenamePattern=RelaxCat_${capturetime:d"yyyyMMdd_HHmmss_fffffff"}_${NUM}
        OutputFileFormat=png
        OutputFileAllowOverwrite=False
        OutputFileCopyPathToClipboard=False
        OutputFilePromptQuality=False
        CaptureWindowsInteractive=False
        PlayCameraSound=False
        ShowTrayNotification=False
        UpdateCheckInterval=0
        """;

    public async ValueTask DisposeAsync()
    {
        stopping = true;
        var process = owned; owned = null;
        if (process is null) return;
        try
        {
            if (!process.HasExited)
            {
                // Never send the global --exit IPC command: it could reach an unrelated replacement instance.
                if (process.CloseMainWindow())
                {
                    using var timeout = new CancellationTokenSource(TimeSpan.FromMilliseconds(750));
                    try { await process.WaitForExitAsync(timeout.Token); } catch (OperationCanceledException) { }
                }
            }
            if (job is not null) await job.StopAsync();
            else if (!process.HasExited) process.Kill(entireProcessTree: true);
            if (!process.HasExited) await process.WaitForExitAsync();
        }
        finally { job?.Dispose(); job = null; process.Dispose(); }
    }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetNamedPipeServerProcessId(SafePipeHandle pipe, out uint processId);
}
