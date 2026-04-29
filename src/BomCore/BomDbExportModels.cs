using System.Text.Json.Serialization;

namespace BomCore;

public sealed record BomDbExportInput
{
    public string? AssemblyPath { get; init; }

    public IReadOnlyDictionary<string, string?> AssemblyCustomProperties { get; init; } =
        new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    public string? ProfilePath { get; init; }

    public BomProfile Profile { get; init; } = new();

    public BomResult Result { get; init; } = new();
}

public sealed record BomDbImportFile
{
    public const string CurrentContractVersion = "bompipe-bomdb.v1";

    [JsonPropertyName("contract_version")]
    public string ContractVersion { get; init; } = CurrentContractVersion;

    [JsonPropertyName("project")]
    public string? Project { get; init; }

    [JsonPropertyName("project_name")]
    public string? ProjectName { get; init; }

    [JsonPropertyName("assembly_path")]
    public string? AssemblyPath { get; init; }

    [JsonPropertyName("rows")]
    public IReadOnlyList<BomDbImportRow> Rows { get; init; } = [];
}

public sealed record BomDbImportRow
{
    [JsonPropertyName("file_path")]
    public string? FilePath { get; init; }

    [JsonPropertyName("configuration_name")]
    public string ConfigurationName { get; init; } = "Default";

    [JsonPropertyName("quantity")]
    public decimal Quantity { get; init; }

    [JsonPropertyName("component_name")]
    public string? ComponentName { get; init; }

    [JsonPropertyName("part_number")]
    public string? PartNumber { get; init; }

    [JsonPropertyName("description")]
    public string? Description { get; init; }

    [JsonPropertyName("material")]
    public string? Material { get; init; }

    [JsonPropertyName("item_number")]
    public string? ItemNumber { get; init; }

    [JsonPropertyName("custom_properties_json")]
    public IReadOnlyDictionary<string, string?> CustomPropertiesJson { get; init; } =
        new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
}
