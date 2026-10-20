using Avalonia;

namespace RelaxCat.App;

internal static class Program
{
    [STAThread]
    public static void Main(string[] args) => AppBuilder.Configure<DesktopApp>().UsePlatformDetect().StartWithClassicDesktopLifetime(args);
}
