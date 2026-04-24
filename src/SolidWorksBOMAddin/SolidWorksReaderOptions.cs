namespace SolidWorksBOMAddin;

public sealed record SolidWorksReaderOptions
{
    public bool IncludeSuppressedComponents { get; init; }

    public bool ResolveLightweightComponents { get; init; }
}
