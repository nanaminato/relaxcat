using System.Collections.Concurrent;
using RelaxCat.Core.Capture;

namespace RelaxCat.Core.Tests;

public class CaptureShortcutTests
{
    [Theory]
    [InlineData("PrintScreen", 44, ShortcutModifiers.None, "PrintScreen")]
    [InlineData("Ctrl+Alt+F8", 119, ShortcutModifiers.Control | ShortcutModifiers.Alt, "Ctrl + Alt + F8")]
    [InlineData("Shift + F24", 135, ShortcutModifiers.Shift, "Shift + F24")]
    [InlineData("Win+K", 75, ShortcutModifiers.Windows, "Win + K")]
    public void Parses_modifiers_and_virtual_keys(string text, int key, ShortcutModifiers modifiers, string formatted)
    {
        var result = CaptureShortcut.Parse(text);
        Assert.Equal(key, result.VirtualKey); Assert.Equal(modifiers, result.Modifiers); Assert.Equal(formatted, result.ToGreenshotString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("Ctrl++F8")]
    [InlineData("Ctrl+Ctrl+F8")]
    [InlineData("F25")]
    [InlineData("Unknown+F1")]
    public void Rejects_invalid_shortcuts(string text) => Assert.Throws<FormatException>(() => CaptureShortcut.Parse(text));
}

public class CaptureFolderMonitorTests
{
    private sealed class Folder : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RelaxCat-CaptureTests-" + Guid.NewGuid().ToString("N"));
        public Folder() => Directory.CreateDirectory(Path);
        public string File(string name) => System.IO.Path.Combine(Path, name);
        public void Dispose() => Directory.Delete(Path, true);
    }
    private static async Task Until(Func<bool> done)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        while (!done()) await Task.Delay(20, timeout.Token);
    }

    [Fact]
    public async Task Waits_until_file_is_stable_and_writer_is_closed_then_imports_once()
    {
        using var folder = new Folder();
        await File.WriteAllTextAsync(folder.File("existing.png"), "baseline");
        var imported = new ConcurrentQueue<string>();
        await using var monitor = new CaptureFolderMonitor(folder.Path, (path, _) => { imported.Enqueue(path); return Task.FromResult(true); },
            settle: TimeSpan.FromMilliseconds(80), poll: TimeSpan.FromMilliseconds(20));
        var path = folder.File("capture.png");
        await using (var writer = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            await writer.WriteAsync(new byte[] { 1, 2 }); await writer.FlushAsync();
            await Task.Delay(220); Assert.Empty(imported);
            await writer.WriteAsync(new byte[] { 3, 4 }); await writer.FlushAsync();
        }
        await Until(() => imported.Count == 1);
        await Task.Delay(300);
        Assert.Equal(new[] { path }, imported.ToArray());
    }

    [Fact]
    public async Task Watches_atomic_rename_and_ignores_unsupported_files()
    {
        using var folder = new Folder(); var imported = new ConcurrentQueue<string>();
        await using var monitor = new CaptureFolderMonitor(folder.Path, (path, _) => { imported.Enqueue(path); return Task.FromResult(true); },
            settle: TimeSpan.FromMilliseconds(50), poll: TimeSpan.FromMilliseconds(20));
        await File.WriteAllTextAsync(folder.File("ignored.txt"), "ignored");
        var temporary = folder.File("capture.tmp"); await File.WriteAllTextAsync(temporary, "data");
        var final = folder.File("capture.JPG"); File.Move(temporary, final);
        await Until(() => imported.Count == 1); Assert.Equal(final, Assert.Single(imported));
    }

    [Fact]
    public async Task A_new_file_version_can_be_imported_and_failed_import_is_retried()
    {
        using var folder = new Folder(); var attempts = 0; var imported = 0;
        await using var monitor = new CaptureFolderMonitor(folder.Path, (_, _) =>
        {
            if (Interlocked.Increment(ref attempts) == 1) return Task.FromResult(false);
            Interlocked.Increment(ref imported); return Task.FromResult(true);
        }, settle: TimeSpan.FromMilliseconds(50), poll: TimeSpan.FromMilliseconds(20));
        var path = folder.File("version.png"); await File.WriteAllTextAsync(path, "first");
        await Until(() => Volatile.Read(ref imported) == 1);
        await File.WriteAllTextAsync(path, "second version");
        await Until(() => Volatile.Read(ref imported) == 2);
        Assert.Equal(3, Volatile.Read(ref attempts));
    }

    [Fact]
    public async Task Stop_cancels_an_inflight_import_and_no_more_files_are_processed()
    {
        using var folder = new Folder(); var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var monitor = new CaptureFolderMonitor(folder.Path, async (_, token) =>
        { entered.TrySetResult(); await Task.Delay(Timeout.Infinite, token); return true; },
            settle: TimeSpan.FromMilliseconds(30), poll: TimeSpan.FromMilliseconds(10));
        await File.WriteAllTextAsync(folder.File("capture.png"), "data");
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await monitor.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(2));
        await File.WriteAllTextAsync(folder.File("after.png"), "data");
    }
}
