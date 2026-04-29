namespace BomCore;

public static class BomExportFormats
{
    public const string Csv = "csv";
    public const string Xlsx = "xlsx";
    public const string BomDbJson = "bomdb-json";

    public static IReadOnlyList<string> Supported { get; } = [Csv, Xlsx, BomDbJson];

    public static string Normalize(string format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            throw new ArgumentException($"Format must be one of: {string.Join(", ", Supported)}.", nameof(format));
        }

        return format.Trim().ToLowerInvariant() switch
        {
            Csv => Csv,
            Xlsx => Xlsx,
            BomDbJson => BomDbJson,
            _ => throw new ArgumentException($"Format must be one of: {string.Join(", ", Supported)}.", nameof(format)),
        };
    }

    public static string GetDefaultExtension(string format)
    {
        return Normalize(format) switch
        {
            Csv => ".csv",
            Xlsx => ".xlsx",
            BomDbJson => ".json",
            _ => throw new InvalidOperationException("Unsupported BOM export format."),
        };
    }

    public static string GetDefaultFileSuffix(string format)
    {
        return Normalize(format) switch
        {
            Csv => ".bom.csv",
            Xlsx => ".bom.xlsx",
            BomDbJson => ".bomdb.json",
            _ => throw new InvalidOperationException("Unsupported BOM export format."),
        };
    }

    public static string GetDisplayName(string format)
    {
        return Normalize(format) switch
        {
            Csv => "CSV BOM",
            Xlsx => "Excel BOM",
            BomDbJson => "BOMDB import JSON",
            _ => throw new InvalidOperationException("Unsupported BOM export format."),
        };
    }

    public static bool TryGetFormatFromPath(string? path, out string format)
    {
        format = string.Empty;
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        if (path.EndsWith(".bomdb.json", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
        {
            format = BomDbJson;
            return true;
        }

        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            format = Xlsx;
            return true;
        }

        if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            format = Csv;
            return true;
        }

        return false;
    }
}
