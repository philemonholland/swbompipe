using System.Text.Json;

namespace BomCore;

public static class BomProfileSerializer
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
    };

    public static BomProfile Deserialize(string json)
    {
        return JsonSerializer.Deserialize<BomProfile>(json, SerializerOptions) ?? new BomProfile();
    }

    public static async Task<BomProfile> DeserializeAsync(Stream input, CancellationToken cancellationToken = default)
    {
        return await JsonSerializer.DeserializeAsync<BomProfile>(input, SerializerOptions, cancellationToken)
            .ConfigureAwait(false)
            ?? new BomProfile();
    }

    public static string Serialize(BomProfile profile)
    {
        return JsonSerializer.Serialize(profile, SerializerOptions);
    }

    public static IReadOnlyList<BomDiagnostic> Validate(BomProfile profile)
    {
        var diagnostics = new List<BomDiagnostic>();
        var enabledColumns = profile.PipeColumns
            .Where(column => column.Enabled)
            .ToList();

        var duplicateDisplayNames = enabledColumns
            .Select(column => column.DisplayName)
            .Concat(profile.AccessoryRules.Select(rule => rule.DisplayName))
            .Where(displayName => !string.IsNullOrWhiteSpace(displayName))
            .GroupBy(displayName => displayName, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1);

        diagnostics.AddRange(
            duplicateDisplayNames.Select(group => new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Error,
                Code = "duplicate-display-name",
                Message = $"Duplicate display name '{group.Key}' found in the BOM profile.",
            }));

        foreach (var requiredProperty in KnownPropertyNames.PipeRequiredProperties)
        {
            if (enabledColumns.All(column => !string.Equals(column.SourceProperty, requiredProperty, StringComparison.OrdinalIgnoreCase)))
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Error,
                    Code = "missing-pipe-field",
                    Message = $"Required pipe field '{requiredProperty}' is missing from enabled pipe columns.",
                    PropertyName = requiredProperty,
                });
            }
        }

        var pipeLengthColumn = enabledColumns.FirstOrDefault(
            column => string.Equals(column.SourceProperty, KnownPropertyNames.PipeLength, StringComparison.OrdinalIgnoreCase));

        if (pipeLengthColumn is not null && !pipeLengthColumn.GroupBy)
        {
            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Error,
                Code = "pipe-length-not-grouped",
                Message = "PipeLength must participate in grouping so equal BOM codes with different lengths stay separate.",
                PropertyName = KnownPropertyNames.PipeLength,
            });
        }

        return diagnostics;
    }
}
