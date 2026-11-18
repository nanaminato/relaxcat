using System.IO.Compression;
using System.Security.Cryptography;

namespace RelaxCat.App.Capture;

public static class GreenshotBundle
{
    public const string Version = "1.4.371-g78211b16d9";
    public const string Sha256 = "b915f7bc8dfde4cdfa271f2731bdc1db9c262a3723e6898efe5c8b3e6385103c";
    private const string VendorFixedSettings = ";In this file you should add your fixed settings";
    private const string WelcomeSetting = "[Core]\r\nWelcomeShown=True\r\n";
    private static readonly SemaphoreSlim installGate = new(1);
#if DEBUG
    public static string DefaultStateRoot => Path.Combine(AppContext.BaseDirectory, "RelaxCat.DebugData");
#else
    public static string DefaultStateRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RelaxCat");
#endif

    public static async Task<string> EnsureAsync(string stateRoot, CancellationToken token)
    {
        await installGate.WaitAsync(token);
        try
        {
            return await Task.Run(() =>
            {
                if (!Environment.Is64BitOperatingSystem) throw new PlatformNotSupportedException("Bundled Greenshot requires 64-bit Windows.");
                var archive = Path.Combine(AppContext.BaseDirectory, "Providers", "Greenshot", "portable.zip");
                using (var stream = File.OpenRead(archive))
                    if (!Convert.ToHexString(SHA256.HashData(stream)).Equals(Sha256, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("Greenshot portable archive checksum mismatch.");
                token.ThrowIfCancellationRequested();
                var folder = Path.Combine(Path.GetFullPath(stateRoot), "Providers", "Greenshot", Version);
                var executable = Path.Combine(folder, "Greenshot.exe");
                var marker = Path.Combine(folder, ".relaxcat-ready");
                if (!File.Exists(marker) || !File.Exists(executable))
                {
                    Directory.CreateDirectory(folder);
                    ZipFile.ExtractToDirectory(archive, folder, true);
                    File.WriteAllText(marker, Sha256);
                }
                // Greenshot reads fixed settings beside the executable, even with --ini-directory.
                // Leave the user's profile and any customized fixed settings intact.
                var fixedSettings = Path.Combine(folder, "greenshot-fixed.ini");
                try
                {
                    if (!File.Exists(fixedSettings) || File.ReadAllText(fixedSettings).Trim() == VendorFixedSettings)
                        File.WriteAllText(fixedSettings, WelcomeSetting);
                }
                catch (UnauthorizedAccessException) { /* The optional welcome override must not block capture. */ }
                catch (IOException) { /* A busy fixed settings file can be retried at the next start. */ }
                token.ThrowIfCancellationRequested();
                return executable;
            }, token);
        }
        finally { installGate.Release(); }
    }
}
