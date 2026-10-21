using System.Diagnostics;
using RelaxCat.Core.Import;

namespace RelaxCat.Core.Capture;

/// <summary>Watches new/changed image versions; existing files form the baseline, not an import backlog.</summary>
public sealed class CaptureFolderMonitor : IAsyncDisposable
{
    private readonly string folder;
    private readonly Func<string, CancellationToken, Task<bool>> import;
    private readonly Action<Exception>? reportError;
    private readonly TimeSpan settle, poll;
    private readonly CancellationTokenSource lifetime = new();
    private readonly FileSystemWatcher watcher;
    private readonly Dictionary<string, Candidate> files = new(StringComparer.OrdinalIgnoreCase);
    private readonly Task worker;
    private int rescan = 1;
    private sealed record Version(long Length, long Modified);
    private sealed class Candidate(Version version, bool baseline)
    {
        public Version Observed = version;
        public Version? Completed = baseline ? version : null;
        public long StableSince = Stopwatch.GetTimestamp();
        public long RetryAt;
        public int Failures;
    }

    public CaptureFolderMonitor(string folder, Func<string, CancellationToken, Task<bool>> import,
        Action<Exception>? reportError = null, TimeSpan? settle = null, TimeSpan? poll = null)
    {
        this.folder = Path.GetFullPath(folder); this.import = import; this.reportError = reportError;
        this.settle = settle ?? TimeSpan.FromMilliseconds(750); this.poll = poll ?? TimeSpan.FromMilliseconds(150);
        if (this.settle < TimeSpan.Zero || this.poll <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(settle));
        Directory.CreateDirectory(this.folder);
        Scan(true);
        watcher = new FileSystemWatcher(this.folder)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.Size | NotifyFilters.LastWrite,
            IncludeSubdirectories = false
        };
        watcher.Created += Changed; watcher.Changed += Changed; watcher.Renamed += Changed;
        watcher.Deleted += Changed;
        watcher.Error += (_, e) => { Interlocked.Exchange(ref rescan, 1); Report(e.GetException()); };
        watcher.EnableRaisingEvents = true;
        worker = Task.Run(RunAsync);
    }

    public static bool Supports(string path) => ImageFileFormats.Supports(path);
    private void Changed(object sender, FileSystemEventArgs e) => Interlocked.Exchange(ref rescan, 1);
    private static Version ReadVersion(string path)
    {
        var info = new FileInfo(path); return new Version(info.Length, info.LastWriteTimeUtc.Ticks);
    }
    private void Scan(bool baseline = false)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in Directory.EnumerateFiles(folder))
        {
            if (!Supports(path)) continue;
            seen.Add(path);
            try
            {
                var version = ReadVersion(path);
                if (!files.TryGetValue(path, out var candidate)) files[path] = new Candidate(version, baseline);
                else if (candidate.Observed != version)
                {
                    candidate.Observed = version; candidate.StableSince = Stopwatch.GetTimestamp();
                    candidate.Failures = 0; candidate.RetryAt = 0;
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        foreach (var deleted in files.Keys.Where(path => !seen.Contains(path)).ToArray()) files.Remove(deleted);
    }

    private async Task RunAsync()
    {
        var token = lifetime.Token; var scanAt = 0L;
        try
        {
            while (!token.IsCancellationRequested)
            {
                // Periodic full reconciliation also recovers missed/overflowed notifications.
                if (Interlocked.Exchange(ref rescan, 0) != 0 || Stopwatch.GetTimestamp() >= scanAt)
                {
                    try { Scan(); } catch (IOException error) { Report(error); } catch (UnauthorizedAccessException error) { Report(error); }
                    scanAt = Stopwatch.GetTimestamp() + Stopwatch.Frequency;
                }
                foreach (var entry in files.ToArray())
                {
                    token.ThrowIfCancellationRequested();
                    var candidate = entry.Value;
                    if (candidate.Observed == candidate.Completed || Stopwatch.GetElapsedTime(candidate.StableSince) < settle || Stopwatch.GetTimestamp() < candidate.RetryAt) continue;
                    try
                    {
                        using (var stream = new FileStream(entry.Key, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            if (stream.Length == 0 || ReadVersion(entry.Key) != candidate.Observed) { Interlocked.Exchange(ref rescan, 1); continue; }
                        }
                        if (await import(entry.Key, token).ConfigureAwait(false)) candidate.Completed = candidate.Observed;
                        else
                        {
                            candidate.RetryAt = Stopwatch.GetTimestamp() + Stopwatch.Frequency;
                            if (++candidate.Failures >= 5)
                            {
                                candidate.Completed = candidate.Observed;
                                Report(new IOException($"Image could not be imported after retries: {entry.Key}"));
                            }
                        }
                    }
                    catch (IOException) { /* Still being written/renamed; retry without marking it imported. */ }
                    catch (UnauthorizedAccessException) { }
                }
                await Task.Delay(poll, token).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { Report(error); }
    }

    private void Report(Exception error) { try { reportError?.Invoke(error); } catch { /* A diagnostic callback must not stop monitoring. */ } }
    public async ValueTask DisposeAsync()
    {
        lifetime.Cancel(); watcher.Dispose();
        await worker.ConfigureAwait(false); lifetime.Dispose();
    }
}
