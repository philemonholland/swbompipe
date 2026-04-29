using System.Text.Json;

namespace BomCore.Tests;

public sealed class BomDbExportTests
{
    [Fact]
    public void Create_BuildsContractRowsFromRepresentativeBomResult()
    {
        var profile = TestData.CreateDefaultProfile();
        var result = new BomGenerator().Generate(
        [
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", quantity: 2m, numGaskets: "2"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "12", quantity: 1m, numGaskets: "2"),
            TestData.CreateClassifiedComponent("fitting-1", KnownBomSections.Fittings, "FT-100", "90 Elbow", quantity: 4m),
        ],
        profile);

        var payload = new BomDbExportService().Create(
            new BomDbExportInput
                {
                    AssemblyPath = @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
                    AssemblyCustomProperties = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Project"] = "PRJ-2026-001",
                        ["Project Name"] = "Test Project",
                    },
                    ProfilePath = @"C:\codebase\bompipe\profiles\default.pipebom.json",
                    Profile = profile,
                    Result = result,
            });

        Assert.Equal(BomDbImportFile.CurrentContractVersion, payload.ContractVersion);
        Assert.Equal("PRJ-2026-001", payload.Project);
        Assert.Equal("Test Project", payload.ProjectName);
        Assert.Equal(@"D:\AFCA\projects\test_project_2\Test_Project.SLDASM", payload.AssemblyPath);
        Assert.Equal(3, payload.Rows.Count);

        var pipeRow = Assert.Single(payload.Rows.Where(row => row.PartNumber == "B-100"));
        Assert.Equal("Default", pipeRow.ConfigurationName);
        Assert.Equal(3m, pipeRow.Quantity);
        Assert.Equal("B-100", pipeRow.ComponentName);
        Assert.Equal("Line A", pipeRow.Description);
        Assert.Equal("Pipes", pipeRow.CustomPropertiesJson["bom_section"]);
        Assert.Equal("pipe_cut", pipeRow.CustomPropertiesJson["row_type"]);
        Assert.Equal("12", pipeRow.CustomPropertiesJson["Cut Length"]);
        Assert.Equal("Spec A", pipeRow.CustomPropertiesJson["Specification"]);

        var fittingRow = Assert.Single(payload.Rows.Where(row => row.PartNumber == "FT-100"));
        Assert.Equal(4m, fittingRow.Quantity);
        Assert.Equal("90 Elbow", fittingRow.Description);
        Assert.Equal("Fittings", fittingRow.CustomPropertiesJson["bom_section"]);

        var accessoryRow = Assert.Single(payload.Rows.Where(row => row.Description == "Gaskets"));
        Assert.Null(accessoryRow.PartNumber);
        Assert.Equal("Gaskets", accessoryRow.ComponentName);
        Assert.Equal(6m, accessoryRow.Quantity);
        Assert.Equal("Other Accessories", accessoryRow.CustomPropertiesJson["bom_section"]);
        Assert.Equal("accessory", accessoryRow.CustomPropertiesJson["row_type"]);
    }

    [Fact]
    public void Create_RejectsRowsWithNonPositiveQuantity()
    {
        var result = new BomResult
        {
            Rows =
            [
                new BomRow
                {
                    Section = KnownBomSections.Other,
                    RowType = BomRowType.SectionItem,
                    Quantity = 0m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        [KnownPropertyNames.Description] = "Invalid row",
                    },
                },
            ],
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new BomDbExportService().Create(
                new BomDbExportInput
                {
                    Profile = TestData.CreateDefaultProfile(),
                    Result = result,
                }));

        Assert.Contains("positive quantity", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Create_AcceptsCommonProjectMetadataAliases()
    {
        var payload = new BomDbExportService().Create(
            new BomDbExportInput
            {
                AssemblyCustomProperties = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Project Number"] = "PRJ-ALIAS-001",
                    ["Project Description"] = "Alias project",
                },
                Profile = TestData.CreateDefaultProfile(),
                Result = new BomResult(),
            });

        Assert.Equal("PRJ-ALIAS-001", payload.Project);
        Assert.Equal("Alias project", payload.ProjectName);
    }

    [Fact]
    public void Export_WritesSnakeCaseContractPayload()
    {
        var payload = new BomDbImportFile
        {
            Project = "PRJ-2026-001",
            ProjectName = "Test Project",
            AssemblyPath = @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
            Rows =
            [
                new BomDbImportRow
                {
                    ComponentName = "B-100",
                    ConfigurationName = "Default",
                    Quantity = 2m,
                    PartNumber = "B-100",
                    Description = "Line A",
                    CustomPropertiesJson = new Dictionary<string, string?>
                    {
                        ["bom_section"] = KnownBomSections.Pipes,
                        ["row_type"] = "pipe_cut",
                        ["BOM Code"] = "B-100",
                        ["Cut Length"] = "12",
                    },
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();

        new BomDbJsonExporter().Export(payload, stream);

        stream.Position = 0;
        using var document = JsonDocument.Parse(stream);
        var root = document.RootElement;
        var row = root.GetProperty("rows")[0];

        Assert.Equal(BomDbImportFile.CurrentContractVersion, root.GetProperty("contract_version").GetString());
        Assert.Equal("PRJ-2026-001", root.GetProperty("project").GetString());
        Assert.Equal("Test Project", root.GetProperty("project_name").GetString());
        Assert.Equal(@"D:\AFCA\projects\test_project_2\Test_Project.SLDASM", root.GetProperty("assembly_path").GetString());
        Assert.Equal("B-100", row.GetProperty("component_name").GetString());
        Assert.Equal("Default", row.GetProperty("configuration_name").GetString());
        Assert.Equal(2m, row.GetProperty("quantity").GetDecimal());
        Assert.Equal("B-100", row.GetProperty("part_number").GetString());
        Assert.Equal("Line A", row.GetProperty("description").GetString());
        Assert.Equal("12", row.GetProperty("custom_properties_json").GetProperty("Cut Length").GetString());
        Assert.False(root.TryGetProperty("schema", out _));
        Assert.False(root.TryGetProperty("schemaVersion", out _));
        Assert.False(root.TryGetProperty("metadata", out _));
        Assert.False(root.TryGetProperty("sections", out _));
    }

    [Fact]
    public void Export_RoundTripsRepresentativeOutputAgainstBomDbImportExpectations()
    {
        var profile = TestData.CreateDefaultProfile();
        var result = new BomGenerator().Generate(
        [
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", quantity: 2m, numGaskets: "2"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "12", quantity: 1m, numGaskets: "2"),
            TestData.CreateClassifiedComponent("fitting-1", KnownBomSections.Fittings, "FT-100", "90 Elbow", quantity: 4m),
        ],
        profile);

        var payload = new BomDbExportService().Create(
            new BomDbExportInput
            {
                Profile = profile,
                Result = result,
            });

        using var stream = TestData.CreateWritableStream();
        new BomDbJsonExporter().Export(payload, stream);

        stream.Position = 0;
        using var document = JsonDocument.Parse(stream);
        var importedRows = document.RootElement
            .GetProperty("rows")
            .EnumerateArray()
            .Select(ImportRow)
            .ToList();

        Assert.Equal(3, importedRows.Count);

        var pipeRow = Assert.Single(importedRows.Where(row => row.CustomProperties["bom_section"] == KnownBomSections.Pipes));
        Assert.Equal(3m, pipeRow.Quantity);
        Assert.Equal("Default", pipeRow.ConfigurationName);
        Assert.Equal("B-100", pipeRow.PartNumber);
        Assert.Equal("Line A", pipeRow.Description);
        Assert.Equal("12", pipeRow.CustomProperties["Cut Length"]);

        var fittingRow = Assert.Single(importedRows.Where(row => row.PartNumber == "FT-100"));
        Assert.Equal("90 Elbow", fittingRow.Description);
        Assert.Equal(4m, fittingRow.Quantity);

        var accessoryRow = Assert.Single(importedRows.Where(row => row.CustomProperties["row_type"] == "accessory"));
        Assert.Equal("Gaskets", accessoryRow.ComponentName);
        Assert.Equal("Gaskets", accessoryRow.PartNumber);
        Assert.Equal(6m, accessoryRow.Quantity);
    }

    [Fact]
    public void Export_RoundTripsPartNumberFromCustomPropertiesFallback()
    {
        var payload = new BomDbImportFile
        {
            Rows =
            [
                new BomDbImportRow
                {
                    ComponentName = "fallback-component.sldprt",
                    ConfigurationName = string.Empty,
                    Quantity = 1m,
                    CustomPropertiesJson = new Dictionary<string, string?>
                    {
                        ["Part Number"] = "PN-42",
                    },
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();
        new BomDbJsonExporter().Export(payload, stream);

        stream.Position = 0;
        using var document = JsonDocument.Parse(stream);
        var importedRow = ImportRow(document.RootElement.GetProperty("rows")[0]);

        Assert.Equal("Default", importedRow.ConfigurationName);
        Assert.Equal("PN-42", importedRow.PartNumber);
    }

    private static ImportedRow ImportRow(JsonElement element)
    {
        var quantity = element.GetProperty("quantity").GetDecimal();
        Assert.True(quantity > 0m);

        var filePath = GetOptionalString(element, "file_path");
        var componentName = GetOptionalString(element, "component_name");
        Assert.True(!string.IsNullOrWhiteSpace(filePath) || !string.IsNullOrWhiteSpace(componentName));

        var configurationName = GetOptionalString(element, "configuration_name");
        configurationName = string.IsNullOrWhiteSpace(configurationName) ? "Default" : configurationName;

        var customProperties = ReadCustomProperties(element);
        var partNumber = FirstNonBlank(
            GetOptionalString(element, "part_number"),
            GetCustomProperty(customProperties, "part_number"),
            GetCustomProperty(customProperties, "Part Number"),
            GetCustomProperty(customProperties, "PartNumber"),
            GetFileNameStem(componentName));

        return new ImportedRow(
            Quantity: quantity,
            ConfigurationName: configurationName,
            ComponentName: componentName,
            PartNumber: partNumber,
            Description: GetOptionalString(element, "description"),
            CustomProperties: customProperties);
    }

    private static IReadOnlyDictionary<string, string?> ReadCustomProperties(JsonElement element)
    {
        if (!element.TryGetProperty("custom_properties_json", out var customPropertiesElement))
        {
            return new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        }

        Assert.Equal(JsonValueKind.Object, customPropertiesElement.ValueKind);

        var properties = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in customPropertiesElement.EnumerateObject())
        {
            properties[property.Name] = property.Value.ValueKind switch
            {
                JsonValueKind.Null => null,
                JsonValueKind.String => property.Value.GetString(),
                _ => property.Value.ToString(),
            };
        }

        return properties;
    }

    private static string? GetCustomProperty(IReadOnlyDictionary<string, string?> values, string key)
    {
        return values.TryGetValue(key, out var value) ? value : null;
    }

    private static string? GetOptionalString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property))
        {
            return null;
        }

        return property.ValueKind == JsonValueKind.Null ? null : property.GetString();
    }

    private static string? GetFileNameStem(string? componentName)
    {
        if (string.IsNullOrWhiteSpace(componentName))
        {
            return null;
        }

        return Path.GetFileNameWithoutExtension(componentName.Trim());
    }

    private static string? FirstNonBlank(params string?[] candidates)
    {
        return candidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate))?.Trim();
    }

    private sealed record ImportedRow(
        decimal Quantity,
        string ConfigurationName,
        string? ComponentName,
        string? PartNumber,
        string? Description,
        IReadOnlyDictionary<string, string?> CustomProperties);
}
