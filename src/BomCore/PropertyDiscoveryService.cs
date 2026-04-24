namespace BomCore;

public sealed class PropertyDiscoveryService
{
    public PropertyDiscoveryResult DiscoverFromComponents(IEnumerable<ComponentRecord> components)
    {
        var propertyNames = components
            .SelectMany(component => component.Properties.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(propertyName => propertyName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var ignoredProperties = propertyNames
            .Where(propertyName => KnownPropertyNames.DefaultIgnoredProperties.Contains(propertyName, StringComparer.OrdinalIgnoreCase))
            .ToList();

        var suggestedProfile = SuggestDefaultProfile(components);

        return new PropertyDiscoveryResult
        {
            DiscoveredProperties = propertyNames,
            SuggestedColumns = suggestedProfile.PipeColumns,
            IgnoredProperties = ignoredProperties,
            Diagnostics = [],
        };
    }

    public BomProfile SuggestDefaultProfile(IEnumerable<ComponentRecord> components)
    {
        var propertyNames = components
            .SelectMany(component => component.Properties.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var pipeColumns = new List<BomColumnRule>();
        var order = 1;

        AddSuggestedColumn(propertyNames, pipeColumns, KnownPropertyNames.Bom, "BOM Code", groupBy: true, ref order);
        AddSuggestedColumn(propertyNames, pipeColumns, KnownPropertyNames.PipeIdentifier, "Pipe Description", groupBy: true, ref order);
        AddSuggestedColumn(propertyNames, pipeColumns, KnownPropertyNames.Specification, "Specification", groupBy: true, ref order);

        if (propertyNames.Contains(KnownPropertyNames.PipeLength))
        {
            pipeColumns.Add(new BomColumnRule
            {
                SourceProperty = KnownPropertyNames.PipeLength,
                DisplayName = "Cut Length",
                Enabled = true,
                GroupBy = true,
                Order = order++,
                Unit = "in",
            });
        }

        var accessoryRules = new List<AccessoryRule>();
        AddAccessoryRule(propertyNames, accessoryRules, KnownPropertyNames.NumGaskets, "Gaskets");
        AddAccessoryRule(propertyNames, accessoryRules, KnownPropertyNames.NumClamps, "Clamps");

        var partClassRules = propertyNames.Contains(KnownPropertyNames.PipeLength)
            ? new List<PartClassRule>
            {
                new()
                {
                    ClassName = "Pipe",
                    DetectWhenPropertyExists = KnownPropertyNames.PipeLength,
                },
            }
            : [];

        return new BomProfile
        {
            ProfileName = "AFCA Pipe BOM",
            Version = 1,
            PartClassRules = partClassRules,
            PipeColumns = pipeColumns,
            AccessoryRules = accessoryRules,
            IgnoredProperties = KnownPropertyNames.DefaultIgnoredProperties.ToList(),
        };
    }

    private static void AddSuggestedColumn(
        IReadOnlySet<string> propertyNames,
        ICollection<BomColumnRule> columns,
        string sourceProperty,
        string displayName,
        bool groupBy,
        ref int order)
    {
        if (!propertyNames.Contains(sourceProperty))
        {
            return;
        }

        columns.Add(new BomColumnRule
        {
            SourceProperty = sourceProperty,
            DisplayName = displayName,
            Enabled = true,
            GroupBy = groupBy,
            Order = order++,
        });
    }

    private static void AddAccessoryRule(
        IReadOnlySet<string> propertyNames,
        ICollection<AccessoryRule> accessoryRules,
        string sourceProperty,
        string displayName)
    {
        if (!propertyNames.Contains(sourceProperty))
        {
            return;
        }

        accessoryRules.Add(new AccessoryRule
        {
            SourceProperty = sourceProperty,
            DisplayName = displayName,
            BomSection = KnownBomSections.PipeAccessories,
        });
    }
}
