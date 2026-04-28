namespace BomCore;

public static class KnownPropertyNames
{
    public const string Bom = "BOM";
    public const string BomDesc = "BOMDesc";
    public const string Description = "Description";
    public const string Class = "Class";
    public const string PrimaryFamily = "Primary Family";
    public const string ComponentType = "Component Type";
    public const string PipeIdentifier = "Pipe Identifier";
    public const string Specification = "Specification";
    public const string PipeLength = "PipeLength";
    public const string TubeLength = "TubeLength";
    public const string WireLength = "WireLength";
    public const string NumGaskets = "NumGaskets";
    public const string NumClamps = "NumClamps";
    public const string BlueGasket = "BlueGasket";
    public const string WhiteGasket = "WhiteGasket";
    public const string BlueFerrule = "BlueFerrule";
    public const string WhiteFerrule = "WhiteFerrule";

    public static readonly IReadOnlyList<string> DefaultIgnoredProperties =
    [
        BlueGasket,
        WhiteGasket,
        BlueFerrule,
        WhiteFerrule,
    ];

    public static readonly IReadOnlyList<string> PipeRequiredProperties =
    [
        Bom,
        PipeIdentifier,
        Specification,
        PipeLength,
    ];

    public static readonly IReadOnlyList<string> PipeCandidateProperties = [];
}

public static class KnownBomSections
{
    public const string Pipes = "Pipes";
    public const string Tubes = "Tubes";
    public const string Wires = "Wires";
    public const string Components = "Components";
    public const string Connections = "Connections";
    public const string Fittings = "Fittings";
    public const string Instruments = "Instruments";
    public const string Systems = "Systems";
    public const string Other = "Other";
    public const string OtherAccessories = "Other Accessories";
    public const string LegacyPipeAccessories = "Pipe Accessories";
    public const string PipeAccessories = OtherAccessories;

    public const string PipeCutList = Pipes;
    public const string OtherComponents = Other;

    public static readonly IReadOnlyList<string> FixedConfigurableSections =
    [
        Pipes,
        Tubes,
        Wires,
    ];

    public static readonly IReadOnlyList<string> PreferredDynamicSections =
    [
        Components,
        Connections,
        Fittings,
        Instruments,
        Systems,
    ];

    public static readonly IReadOnlyList<string> DefaultVisibleConfigurableSections =
    [
        Pipes,
        Tubes,
        Wires,
        Other,
    ];

    public static readonly IReadOnlyList<string> DisplayOrder =
    [
        Pipes,
        Tubes,
        Wires,
        Components,
        Connections,
        Fittings,
        Instruments,
        Systems,
        Other,
        OtherAccessories,
    ];

    public static readonly IReadOnlyList<string> ConfigurableSections =
    [
        Pipes,
        Tubes,
        Wires,
        Components,
        Connections,
        Fittings,
        Instruments,
        Systems,
        Other,
    ];

    public static readonly IReadOnlySet<string> ConfigurableSectionSet =
        ConfigurableSections.ToHashSet(StringComparer.OrdinalIgnoreCase);

    public static bool IsConfigurableSection(string? section)
    {
        if (string.IsNullOrWhiteSpace(section))
        {
            return false;
        }

        var normalizedSection = NormalizeConfigurableSection(section);
        return !string.Equals(normalizedSection, OtherAccessories, StringComparison.OrdinalIgnoreCase);
    }

    public static string NormalizeConfigurableSection(string? section)
    {
        if (string.IsNullOrWhiteSpace(section))
        {
            return Other;
        }

        if (string.Equals(section, LegacyPipeAccessories, StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, PipeAccessories, StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, OtherAccessories, StringComparison.OrdinalIgnoreCase))
        {
            return OtherAccessories;
        }

        return ConfigurableSections
            .Concat(FixedConfigurableSections)
            .FirstOrDefault(knownSection => string.Equals(knownSection, section, StringComparison.OrdinalIgnoreCase))
            ?? section.Trim();
    }

    public static string NormalizeAccessorySection(string? section)
    {
        if (string.IsNullOrWhiteSpace(section)
            || string.Equals(section, LegacyPipeAccessories, StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, PipeAccessories, StringComparison.OrdinalIgnoreCase))
        {
            return OtherAccessories;
        }

        return section.Trim();
    }

    public static bool IsPipeLikeSection(string section)
    {
        return string.Equals(section, Pipes, StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, Tubes, StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, Wires, StringComparison.OrdinalIgnoreCase);
    }

    public static string GetLengthProperty(string section)
    {
        if (string.Equals(section, Tubes, StringComparison.OrdinalIgnoreCase))
        {
            return KnownPropertyNames.TubeLength;
        }

        if (string.Equals(section, Wires, StringComparison.OrdinalIgnoreCase))
        {
            return KnownPropertyNames.WireLength;
        }

        return KnownPropertyNames.PipeLength;
    }

    public static readonly IReadOnlyList<string> ClassMappedSections =
    [
        Pipes,
        Tubes,
        Wires,
        Components,
        Connections,
        Fittings,
        Instruments,
        Systems,
        Other,
    ];

    public static readonly IReadOnlySet<string> ClassMappedSectionSet =
        ClassMappedSections.ToHashSet(StringComparer.OrdinalIgnoreCase);

    public static bool IsClassMappedSection(string? section)
    {
        return !string.IsNullOrWhiteSpace(section) && ClassMappedSectionSet.Contains(section);
    }

    public static bool IsAccessorySection(string? section)
    {
        return string.Equals(NormalizeAccessorySection(section), OtherAccessories, StringComparison.OrdinalIgnoreCase);
    }

    public static IReadOnlyList<string> BuildConfigurableSections(IEnumerable<string>? candidateSections)
    {
        var normalizedSections = (candidateSections ?? [])
            .Select(NormalizeConfigurableSection)
            .Where(section => !string.Equals(section, OtherAccessories, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var dynamicSections = normalizedSections
            .Where(section => !FixedConfigurableSections.Contains(section, StringComparer.OrdinalIgnoreCase)
                && !string.Equals(section, Other, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var orderedDynamicSections = PreferredDynamicSections
            .Where(section => dynamicSections.Contains(section, StringComparer.OrdinalIgnoreCase))
            .Concat(dynamicSections
                .Where(section => !PreferredDynamicSections.Contains(section, StringComparer.OrdinalIgnoreCase))
                .OrderBy(section => section, StringComparer.OrdinalIgnoreCase))
            .ToList();

        return FixedConfigurableSections
            .Concat(orderedDynamicSections)
            .Append(Other)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static IReadOnlyList<string> OrderSections(IEnumerable<string> sections)
    {
        var normalizedSections = (sections ?? [])
            .Select(section => IsAccessorySection(section) ? OtherAccessories : NormalizeConfigurableSection(section))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var orderedSections = BuildConfigurableSections(
            normalizedSections.Where(section => !string.Equals(section, OtherAccessories, StringComparison.OrdinalIgnoreCase)))
            .Where(section => normalizedSections.Contains(section, StringComparer.OrdinalIgnoreCase))
            .ToList();

        if (normalizedSections.Contains(OtherAccessories, StringComparer.OrdinalIgnoreCase))
        {
            orderedSections.Add(OtherAccessories);
        }

        return orderedSections;
    }
}

public static class KnownBomColumnProfiles
{
    public static IReadOnlyList<BomColumnRule> CreateDefaultSectionColumns(string section)
    {
        if (string.Equals(section, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase))
        {
            return
            [
                CreateColumn(KnownPropertyNames.Bom, "BOM Code", 1, groupBy: true),
                CreateColumn(KnownPropertyNames.PipeIdentifier, "Pipe Description", 2, groupBy: true),
                CreateColumn(KnownPropertyNames.ComponentType, "Component Type", 3, groupBy: true),
                CreateColumn(KnownPropertyNames.Specification, "Specification", 4, groupBy: true),
                CreateColumn(KnownPropertyNames.PipeLength, "Cut Length", 5, groupBy: true, unit: "in"),
            ];
        }

        if (string.Equals(section, KnownBomSections.Tubes, StringComparison.OrdinalIgnoreCase))
        {
            return CreateLengthColumns(KnownPropertyNames.TubeLength, "Tube Length");
        }

        if (string.Equals(section, KnownBomSections.Wires, StringComparison.OrdinalIgnoreCase))
        {
            return CreateLengthColumns(KnownPropertyNames.WireLength, "Wire Length");
        }

        return
        [
            CreateColumn(KnownPropertyNames.BomDesc, KnownPropertyNames.BomDesc, 1, groupBy: true),
            CreateColumn(KnownPropertyNames.Description, KnownPropertyNames.Description, 2, groupBy: true),
            CreateColumn(KnownPropertyNames.ComponentType, KnownPropertyNames.ComponentType, 3, groupBy: true),
        ];
    }

    private static IReadOnlyList<BomColumnRule> CreateLengthColumns(string lengthProperty, string lengthHeader)
    {
        return
        [
            CreateColumn(KnownPropertyNames.BomDesc, KnownPropertyNames.BomDesc, 1, groupBy: true),
            CreateColumn(KnownPropertyNames.Description, KnownPropertyNames.Description, 2, groupBy: true),
            CreateColumn(KnownPropertyNames.ComponentType, KnownPropertyNames.ComponentType, 3, groupBy: true),
            CreateColumn(lengthProperty, lengthHeader, 4, groupBy: true, unit: "in"),
        ];
    }

    private static BomColumnRule CreateColumn(
        string sourceProperty,
        string displayName,
        int order,
        bool groupBy,
        string? unit = null)
    {
        return new BomColumnRule
        {
            SourceProperty = sourceProperty,
            DisplayName = displayName,
            Enabled = true,
            GroupBy = groupBy,
            Order = order,
            Unit = unit,
        };
    }
}
