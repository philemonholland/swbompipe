namespace BomCore;

public sealed record BomFileExportRequest
{
    public string Format { get; init; } = BomExportFormats.Csv;

    public string? AssemblyPath { get; init; }

    public IReadOnlyDictionary<string, string?> AssemblyCustomProperties { get; init; } =
        new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    public string? ProfilePath { get; init; }

    public BomProfile Profile { get; init; } = new();

    public BomResult Result { get; init; } = new();
}

public sealed class BomFileExportService
{
    private readonly BomDbExportService _bomDbExportService = new();
    private readonly BomDbJsonExporter _bomDbJsonExporter = new();

    public void Export(BomFileExportRequest request, Stream output)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(output);

        switch (BomExportFormats.Normalize(request.Format))
        {
            case BomExportFormats.Csv:
                new CsvBomExporter().Export(request.Result, output);
                break;
            case BomExportFormats.Xlsx:
                new XlsxBomExporter().Export(request.Result, output);
                break;
            case BomExportFormats.BomDbJson:
                var payload = _bomDbExportService.Create(
                    new BomDbExportInput
                    {
                        AssemblyPath = request.AssemblyPath,
                        AssemblyCustomProperties = request.AssemblyCustomProperties,
                        ProfilePath = request.ProfilePath,
                        Profile = request.Profile,
                        Result = request.Result,
                    });
                _bomDbJsonExporter.Export(payload, output);
                break;
            default:
                throw new InvalidOperationException("Unsupported BOM export format.");
        }
    }
}
