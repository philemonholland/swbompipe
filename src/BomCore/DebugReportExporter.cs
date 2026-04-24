using System.Globalization;
using System.Text;
using System.Text.Json;

namespace BomCore;

public sealed class DebugReportExporter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    public void Export(DebugReport report, Stream output, DebugReportFormat format)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentNullException.ThrowIfNull(output);

        switch (format)
        {
            case DebugReportFormat.Json:
                JsonSerializer.Serialize(output, report, JsonOptions);
                break;
            case DebugReportFormat.Text:
                ExportText(report, output);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported debug report format.");
        }
    }

    public static DebugReportFormat GetFormatFromPath(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var extension = Path.GetExtension(path);
        return extension.ToLowerInvariant() switch
        {
            ".json" => DebugReportFormat.Json,
            ".txt" => DebugReportFormat.Text,
            _ => throw new ArgumentException("Debug report path must end with '.txt' or '.json'.", nameof(path)),
        };
    }

    private static void ExportText(DebugReport report, Stream output)
    {
        using var writer = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true);

        writer.WriteLine("AFCA Piping BOM Debug Report");
        writer.WriteLine($"Assembly Path: {report.AssemblyPath ?? "(not set)"}");
        writer.WriteLine($"Profile Path: {report.ProfilePath ?? "(not set)"}");
        writer.WriteLine($"Components Scanned: {report.ComponentsScanned.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine(
            $"Components Skipped: {FormatOptionalCount(report.ComponentsSkipped)}");
        writer.WriteLine($"Generated BOM Rows: {report.GeneratedBomRowCount.ToString(CultureInfo.InvariantCulture)}");
        writer.WriteLine($"Unique Discovered Properties ({report.DiscoveredProperties.Count.ToString(CultureInfo.InvariantCulture)}):");

        if (report.DiscoveredProperties.Count == 0)
        {
            writer.WriteLine("  (none)");
        }
        else
        {
            foreach (var propertyName in report.DiscoveredProperties)
            {
                writer.WriteLine($"  - {propertyName}");
            }
        }

        writer.WriteLine($"Diagnostics ({report.Diagnostics.Count.ToString(CultureInfo.InvariantCulture)}):");
        if (report.Diagnostics.Count == 0)
        {
            writer.WriteLine("  (none)");
            return;
        }

        foreach (var diagnostic in report.Diagnostics)
        {
            writer.WriteLine($"  - [{diagnostic.Severity}] {diagnostic.Code}: {diagnostic.Message}{FormatDiagnosticContext(diagnostic)}");
        }
    }

    private static string FormatOptionalCount(int? value)
    {
        return value.HasValue
            ? value.Value.ToString(CultureInfo.InvariantCulture)
            : "n/a";
    }

    private static string FormatDiagnosticContext(BomDiagnostic diagnostic)
    {
        var context = new List<string>();

        if (!string.IsNullOrWhiteSpace(diagnostic.ComponentId))
        {
            context.Add($"component={diagnostic.ComponentId}");
        }

        if (!string.IsNullOrWhiteSpace(diagnostic.PropertyName))
        {
            context.Add($"property={diagnostic.PropertyName}");
        }

        return context.Count == 0
            ? string.Empty
            : $" ({string.Join(", ", context)})";
    }
}
