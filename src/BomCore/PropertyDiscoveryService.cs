namespace BomCore;

public sealed class PropertyDiscoveryService
{
    public PropertyDiscoveryResult DiscoverFromComponents(IEnumerable<ComponentRecord> components)
    {
        var componentList = components.ToList();
        var propertyNames = components
            .SelectMany(component => component.Properties.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(propertyName => propertyName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var ignoredProperties = propertyNames
            .Where(propertyName => KnownPropertyNames.DefaultIgnoredProperties.Contains(propertyName, StringComparer.OrdinalIgnoreCase))
            .ToList();

        var discoveredSections = DiscoverSections(componentList);
        var suggestedProfile = SuggestDefaultProfile(componentList);

        return new PropertyDiscoveryResult
        {
            DiscoveredProperties = propertyNames,
            DiscoveredSections = discoveredSections,
            SuggestedColumns = suggestedProfile.GetSectionColumns(KnownBomSections.Pipes),
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

        var discoveredSections = KnownBomSections.BuildConfigurableSections(DiscoverSections(components));

        var sectionProfiles = new List<BomSectionColumnProfile>();

        foreach (var section in discoveredSections)
        {
            var columns = KnownBomColumnProfiles.CreateDefaultSectionColumns(section)
                .Where(column => propertyNames.Contains(column.SourceProperty))
                .ToList();

            if (columns.Count == 0 && !KnownBomSections.IsPipeLikeSection(section))
            {
                columns = KnownBomColumnProfiles.CreateDefaultSectionColumns(section).ToList();
            }

            sectionProfiles.Add(new BomSectionColumnProfile
            {
                Section = section,
                Columns = columns.Count == 0 ? KnownBomColumnProfiles.CreateDefaultSectionColumns(section) : columns,
            });
        }

        var accessoryRules = new List<AccessoryRule>();
        AddAccessoryRule(propertyNames, accessoryRules, KnownPropertyNames.NumGaskets, "Gaskets");
        AddAccessoryRule(propertyNames, accessoryRules, KnownPropertyNames.NumClamps, "Clamps");

        var sectionRules = KnownBomSections.ClassMappedSections
            .Select(section => new BomSectionRule
            {
                SourceProperty = KnownPropertyNames.PrimaryFamily,
                MatchValue = section,
                Section = section,
            })
            .ToList();

        return new BomProfile
        {
            ProfileName = "AFCA Pipe BOM",
            Version = 2,
            PartClassRules =
            [
                new PartClassRule
                {
                    ClassName = "Pipe",
                    DetectWhenPropertyExists = KnownPropertyNames.PipeLength,
                },
            ],
            PipeColumns = sectionProfiles
                .First(profile => string.Equals(profile.Section, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase))
                .Columns,
            SectionColumnProfiles = sectionProfiles,
            SectionRules = sectionRules,
            AccessoryRules = accessoryRules,
            IgnoredProperties = KnownPropertyNames.DefaultIgnoredProperties.ToList(),
        };
    }

    private static IReadOnlyList<string> DiscoverSections(IEnumerable<ComponentRecord> components)
    {
        return components
            .Select(component =>
            {
                var familyValue = component.GetPropertyValue(KnownPropertyNames.PrimaryFamily);
                return string.IsNullOrWhiteSpace(familyValue)
                    ? KnownBomSections.Other
                    : KnownBomSections.NormalizeConfigurableSection(familyValue);
            })
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(section => section, StringComparer.OrdinalIgnoreCase)
            .ToList();
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
