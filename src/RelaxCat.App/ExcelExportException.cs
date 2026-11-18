namespace RelaxCat.App;

internal sealed class ExcelExportException(string stageKey, Exception cause, string? fileName = null)
    : Exception(cause.Message, cause)
{
    internal string Describe(Func<string, string> translate)
    {
        var stage = translate(stageKey);
        if (fileName is not null) stage = string.Format(stage, fileName);
        return $"{stage} (0x{InnerException!.HResult:X8}) — {Message}";
    }
}
