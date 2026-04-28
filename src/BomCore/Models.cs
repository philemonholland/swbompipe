using System.Globalization;

namespace BomCore;

public enum PropertyScope
{
    File,
    Configuration,
    Component,
    CutList,
    Unknown,
}

public enum BomRowType
{
    SectionItem,
    PipeCut,
    Fitting,
    Accessory,
    Other,
}

public enum DiagnosticSeverity
{
    Info,
    Warning,
    Error,
}

public enum DebugReportFormat
{
    Text,
    Json,
}

public sealed record PropertyValue
{
    public string Name { get; init; } = string.Empty;

    public string? RawValue { get; init; }

    public string? EvaluatedValue { get; init; }

    public PropertyScope Scope { get; init; } = PropertyScope.Unknown;

    public string? Source { get; init; }

    public string EffectiveValue => string.IsNullOrWhiteSpace(EvaluatedValue) ? RawValue ?? string.Empty : EvaluatedValue;
}

public sealed record ComponentRecord
{
    public string ComponentId { get; init; } = string.Empty;

    public string? ParentComponentId { get; init; }

    public string? FilePath { get; init; }

    public string ConfigurationName { get; init; } = string.Empty;

    public string ComponentName { get; init; } = string.Empty;

    public decimal Quantity { get; init; } = 1m;

    public bool IsSuppressed { get; init; }

    public bool IsHidden { get; init; }

    public bool IsVirtual { get; init; }

    public bool IsAssembly { get; init; }

    public IReadOnlyDictionary<string, PropertyValue> Properties { get; init; } =
        new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase);

    public bool TryGetProperty(string propertyName, out PropertyValue? propertyValue)
    {
        if (Properties.TryGetValue(propertyName, out var directMatch))
        {
            propertyValue = directMatch;
            return true;
        }

        foreach (var pair in Properties)
        {
            if (string.Equals(pair.Key, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                propertyValue = pair.Value;
                return true;
            }
        }

        propertyValue = null;
        return false;
    }

    public string? GetPropertyValue(string propertyName)
    {
        return TryGetProperty(propertyName, out var propertyValue)
            ? NormalizeValue(propertyValue!.EffectiveValue)
            : null;
    }

    public string GetIdentityFallback()
    {
        if (!string.IsNullOrWhiteSpace(FilePath))
        {
            return $"{FilePath}|{ConfigurationName}";
        }

        return $"virtual:{ParentComponentId ?? "root"}|{ComponentName}|{ConfigurationName}";
    }

    private static string? NormalizeValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim();
    }
}

public sealed record PartClassRule
{
    public string ClassName { get; init; } = string.Empty;

    public string DetectWhenPropertyExists { get; init; } = string.Empty;
}

public sealed record BomColumnRule
{
    public string SourceProperty { get; init; } = string.Empty;

    public string DisplayName { get; init; } = string.Empty;

    public bool Enabled { get; init; }

    public bool GroupBy { get; init; }

    public int Order { get; init; }

    public string? Unit { get; init; }
}

public sealed record BomSectionColumnProfile
{
    public string Section { get; init; } = KnownBomSections.Other;

    public IReadOnlyList<BomColumnRule> Columns { get; init; } = [];
}

public sealed record BomSectionRule
{
    public string SourceProperty { get; init; } = KnownPropertyNames.PrimaryFamily;

    public string MatchValue { get; init; } = string.Empty;

    public string Section { get; init; } = KnownBomSections.Other;
}

public sealed record AccessoryRule
{
    public string SourceProperty { get; init; } = string.Empty;

    public string DisplayName { get; init; } = string.Empty;

    public string BomSection { get; init; } = KnownBomSections.PipeAccessories;
}

public sealed record BomProfile
{
    public string ProfileName { get; init; } = string.Empty;

    public int Version { get; init; } = 1;

    public IReadOnlyList<PartClassRule> PartClassRules { get; init; } = [];

    public IReadOnlyList<BomColumnRule> PipeColumns { get; init; } = [];

    public IReadOnlyList<BomSectionColumnProfile> SectionColumnProfiles { get; init; } = [];

    public IReadOnlyList<BomSectionRule> SectionRules { get; init; } = [];

    public IReadOnlyList<AccessoryRule> AccessoryRules { get; init; } = [];

    public IReadOnlyList<string> IgnoredProperties { get; init; } = [];

    public string GetPipeDetectionProperty()
    {
        return PartClassRules
            .FirstOrDefault(rule => string.Equals(rule.ClassName, "Pipe", StringComparison.OrdinalIgnoreCase))
            ?.DetectWhenPropertyExists
            ?? KnownPropertyNames.PipeLength;
    }

    public IReadOnlyList<BomColumnRule> GetSectionColumns(string section)
    {
        var normalizedSection = KnownBomSections.NormalizeConfigurableSection(section);
        var configuredColumns = SectionColumnProfiles
            .FirstOrDefault(profile => string.Equals(profile.Section, normalizedSection, StringComparison.OrdinalIgnoreCase))
            ?.Columns;

        if (configuredColumns is { Count: > 0 })
        {
            return configuredColumns;
        }

        if (string.Equals(normalizedSection, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase)
            && PipeColumns.Count > 0)
        {
            return PipeColumns;
        }

        return KnownBomColumnProfiles.CreateDefaultSectionColumns(normalizedSection);
    }

    public IReadOnlyList<string> GetConfiguredConfigurableSections()
    {
        return KnownBomSections.BuildConfigurableSections(
            SectionColumnProfiles.Select(profile => profile.Section)
                .Concat(SectionRules.Select(rule => rule.Section))
                .Concat([KnownBomSections.Other]));
    }

    public IReadOnlyList<string> GetEffectiveConfigurableSections(IEnumerable<string>? discoveredSections = null)
    {
        return KnownBomSections.BuildConfigurableSections(
            GetConfiguredConfigurableSections()
                .Concat(discoveredSections ?? []));
    }

    public IReadOnlyList<BomSectionColumnProfile> GetEffectiveSectionColumnProfiles(IEnumerable<string>? discoveredSections = null)
    {
        var profiles = new List<BomSectionColumnProfile>();
        foreach (var section in GetEffectiveConfigurableSections(discoveredSections))
        {
            profiles.Add(new BomSectionColumnProfile
            {
                Section = section,
                Columns = GetSectionColumns(section),
            });
        }

        return profiles;
    }

    public IReadOnlyList<BomSectionRule> GetEffectiveSectionRules()
    {
        if (SectionRules.Count > 0)
        {
            return SectionRules
                .Where(rule => !string.IsNullOrWhiteSpace(rule.SourceProperty)
                    && !string.IsNullOrWhiteSpace(rule.MatchValue)
                    && KnownBomSections.IsConfigurableSection(rule.Section))
                .ToList();
        }

        return KnownBomSections.ClassMappedSections
            .Select(section => new BomSectionRule
            {
                SourceProperty = KnownPropertyNames.PrimaryFamily,
                MatchValue = section,
                Section = section,
            })
            .ToList();
    }
}

public sealed record BomRow
{
    public string Section { get; init; } = KnownBomSections.OtherComponents;

    public BomRowType RowType { get; init; } = BomRowType.Other;

    public IReadOnlyDictionary<string, string> Values { get; init; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public decimal Quantity { get; init; }
}

public sealed record BomDiagnostic
{
    public DiagnosticSeverity Severity { get; init; } = DiagnosticSeverity.Info;

    public string Code { get; init; } = string.Empty;

    public string Message { get; init; } = string.Empty;

    public string? ComponentId { get; init; }

    public string? PropertyName { get; init; }
}

public sealed record BomResult
{
    public IReadOnlyList<BomRow> Rows { get; init; } = [];

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record PropertyDiscoveryResult
{
    public IReadOnlyList<string> DiscoveredProperties { get; init; } = [];

    public IReadOnlyList<string> DiscoveredSections { get; init; } = [];

    public IReadOnlyList<BomColumnRule> SuggestedColumns { get; init; } = [];

    public IReadOnlyList<string> IgnoredProperties { get; init; } = [];

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record ProfileLoadResult
{
    public BomProfile Profile { get; init; } = new();

    public string? SourcePath { get; init; }

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record AssemblyReadResult
{
    public string? AssemblyPath { get; init; }

    public IReadOnlyList<ComponentRecord> Components { get; init; } = [];

    public int ComponentsScanned { get; init; }

    public int? ComponentsSkipped { get; init; }

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record DebugReportInput
{
    public string? AssemblyPath { get; init; }

    public string? ProfilePath { get; init; }

    public IReadOnlyList<ComponentRecord> Components { get; init; } = [];

    public int? ComponentsScanned { get; init; }

    public int? ComponentsSkipped { get; init; }

    public IReadOnlyList<BomRow> Rows { get; init; } = [];

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record DebugReport
{
    public string? AssemblyPath { get; init; }

    public string? ProfilePath { get; init; }

    public int ComponentsScanned { get; init; }

    public int? ComponentsSkipped { get; init; }

    public IReadOnlyList<string> DiscoveredProperties { get; init; } = [];

    public int GeneratedBomRowCount { get; init; }

    public IReadOnlyList<BomDiagnostic> Diagnostics { get; init; } = [];
}

public sealed record ProfileStoreOptions
{
    public string DefaultProfilePath { get; init; } = string.Empty;

    public string? UserProfileDirectory { get; init; }

    public string? CompanyProfileDirectory { get; init; }

    public string DefaultProfileFileName { get; init; } = "default.pipebom.json";
}

internal static class NumericParsing
{
    public static bool TryParseDecimal(string? value, out decimal result)
    {
        return decimal.TryParse(
            value,
            NumberStyles.Number,
            CultureInfo.InvariantCulture,
            out result)
            || decimal.TryParse(
                value,
                NumberStyles.Number,
                CultureInfo.CurrentCulture,
                out result);
    }
}
