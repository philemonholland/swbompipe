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
        var effectiveProfiles = profile.GetEffectiveSectionColumnProfiles();
        var enabledColumns = effectiveProfiles
            .SelectMany(sectionProfile => sectionProfile.Columns.Select(column => new { sectionProfile.Section, Column = column }))
            .Where(entry => entry.Column.Enabled)
            .ToList();

        var duplicateDisplayNames = enabledColumns
            .GroupBy(entry => entry.Section, StringComparer.OrdinalIgnoreCase)
            .SelectMany(sectionGroup => sectionGroup
                .Select(entry => entry.Column.DisplayName)
                .Where(displayName => !string.IsNullOrWhiteSpace(displayName))
                .GroupBy(displayName => displayName, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key))
            .Concat(profile.GetSectionColumns(KnownBomSections.Pipes)
                .Where(column => column.Enabled)
                .Select(column => column.DisplayName)
                .Concat(profile.AccessoryRules.Select(rule => rule.DisplayName))
                .Where(displayName => !string.IsNullOrWhiteSpace(displayName))
                .GroupBy(displayName => displayName, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key))
            .Where(displayName => !string.IsNullOrWhiteSpace(displayName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        diagnostics.AddRange(
            duplicateDisplayNames.Select(displayName => new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Error,
                Code = "duplicate-display-name",
                Message = $"Duplicate display name '{displayName}' found in the BOM profile.",
            }));

        foreach (var requiredProperty in KnownPropertyNames.PipeRequiredProperties)
        {
            var pipeColumns = profile.GetSectionColumns(KnownBomSections.Pipes).Where(column => column.Enabled).ToList();
            if (pipeColumns.All(column => !string.Equals(column.SourceProperty, requiredProperty, StringComparison.OrdinalIgnoreCase)))
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

        foreach (var section in new[] { KnownBomSections.Pipes, KnownBomSections.Tubes, KnownBomSections.Wires })
        {
            var lengthProperty = KnownBomSections.GetLengthProperty(section);
            var lengthColumn = profile.GetSectionColumns(section)
                .Where(column => column.Enabled)
                .FirstOrDefault(column => string.Equals(column.SourceProperty, lengthProperty, StringComparison.OrdinalIgnoreCase));

            if (lengthColumn is not null && !lengthColumn.GroupBy)
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Error,
                    Code = "length-not-grouped",
                    Message = $"{lengthProperty} must participate in grouping so equal BOM codes with different lengths stay separate.",
                    PropertyName = lengthProperty,
                });
            }
        }

        return diagnostics;
    }
}
