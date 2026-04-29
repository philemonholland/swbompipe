namespace BomCore;

public sealed class BomDbExportService
{
    private static readonly IReadOnlyList<string> PartNumberCandidates =
    [
        "part_number",
        "Part Number",
        "PartNumber",
        "BOM Code",
        KnownPropertyNames.Bom,
        KnownPropertyNames.BomDesc,
        "SWbompartno",
    ];

    private static readonly IReadOnlyList<string> DescriptionCandidates =
    [
        "description",
        KnownPropertyNames.Description,
        "Pipe Description",
        KnownPropertyNames.PipeIdentifier,
        KnownPropertyNames.BomDesc,
    ];

    private static readonly IReadOnlyList<string> MaterialCandidates =
    [
        "material",
        "Material",
    ];

    private static readonly IReadOnlyList<string> ItemNumberCandidates =
    [
        "item_number",
        "Item Number",
        "Item No.",
    ];

    private static readonly IReadOnlyList<string> ProjectCandidates =
    [
        "project",
        "Project",
        "Project Number",
        "Project No",
        "Project No.",
        "Project #",
    ];

    private static readonly IReadOnlyList<string> ProjectNameCandidates =
    [
        "project_name",
        "Project Name",
        "ProjectName",
        "Project_Name",
        "Project Description",
    ];

    public BomDbImportFile Create(BomDbExportInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(input.Profile);
        ArgumentNullException.ThrowIfNull(input.Result);

        return new BomDbImportFile
        {
            Project = FindFirstValue(input.AssemblyCustomProperties, ProjectCandidates),
            ProjectName = FindFirstValue(input.AssemblyCustomProperties, ProjectNameCandidates),
            AssemblyPath = NormalizeOptionalText(input.AssemblyPath),
            Rows = (input.Result.Rows ?? [])
                .Select(CreateRow)
                .ToList(),
        };
    }

    private static BomDbImportRow CreateRow(BomRow row)
    {
        if (row.Quantity <= 0m)
        {
            throw new InvalidOperationException(
                $"BOMDB export rows must have a positive quantity. Section '{row.Section}' produced '{row.Quantity}'.");
        }

        var partNumber = FindFirstRowValue(row.Values, PartNumberCandidates);
        var description = FindFirstRowValue(row.Values, DescriptionCandidates);
        var material = FindFirstRowValue(row.Values, MaterialCandidates);
        var itemNumber = FindFirstRowValue(row.Values, ItemNumberCandidates);

        return new BomDbImportRow
        {
            ConfigurationName = "Default",
            Quantity = row.Quantity,
            ComponentName = FirstNonBlank(partNumber, description, NormalizeSection(row.Section)),
            PartNumber = partNumber,
            Description = description,
            Material = material,
            ItemNumber = itemNumber,
            CustomPropertiesJson = BuildCustomProperties(row),
        };
    }

    private static IReadOnlyDictionary<string, string?> BuildCustomProperties(BomRow row)
    {
        var properties = new SortedDictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            ["bom_section"] = NormalizeSection(row.Section),
            ["row_type"] = FormatRowType(row.RowType),
        };

        foreach (var pair in row.Values
                     .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && !string.IsNullOrWhiteSpace(pair.Value))
                     .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
        {
            properties[pair.Key] = pair.Value;
        }

        return properties;
    }

    private static string? FindFirstRowValue(IReadOnlyDictionary<string, string> values, IEnumerable<string> candidates)
    {
        foreach (var candidate in candidates)
        {
            if (TryGetRowValue(values, candidate, out var value))
            {
                return value;
            }
        }

        return null;
    }

    private static bool TryGetRowValue(IReadOnlyDictionary<string, string> values, string key, out string? value)
    {
        if (values.TryGetValue(key, out var directValue) && !string.IsNullOrWhiteSpace(directValue))
        {
            value = directValue;
            return true;
        }

        foreach (var pair in values)
        {
            if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(pair.Value))
            {
                value = pair.Value;
                return true;
            }
        }

        value = null;
        return false;
    }

    private static bool TryGetValue(IReadOnlyDictionary<string, string?> values, string key, out string? value)
    {
        if (values.TryGetValue(key, out var directValue) && !string.IsNullOrWhiteSpace(directValue))
        {
            value = directValue.Trim();
            return true;
        }

        foreach (var pair in values)
        {
            if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(pair.Value))
            {
                value = pair.Value.Trim();
                return true;
            }
        }

        value = null;
        return false;
    }

    private static string? FindFirstValue(IReadOnlyDictionary<string, string?> values, IEnumerable<string> candidates)
    {
        foreach (var candidate in candidates)
        {
            if (TryGetValue(values, candidate, out var value))
            {
                return value;
            }
        }

        return null;
    }

    private static string? NormalizeOptionalText(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    private static string FirstNonBlank(params string?[] candidates)
    {
        return candidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate))?.Trim()
            ?? KnownBomSections.Other;
    }

    private static string NormalizeSection(string? section)
    {
        return KnownBomSections.IsAccessorySection(section)
            ? KnownBomSections.OtherAccessories
            : KnownBomSections.NormalizeConfigurableSection(section);
    }

    private static string FormatRowType(BomRowType rowType)
    {
        return rowType switch
        {
            BomRowType.SectionItem => "section_item",
            BomRowType.PipeCut => "pipe_cut",
            BomRowType.Fitting => "fitting",
            BomRowType.Accessory => "accessory",
            _ => "other",
        };
    }
}
