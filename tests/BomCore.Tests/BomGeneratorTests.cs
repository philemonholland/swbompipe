namespace BomCore.Tests;

public sealed class BomGeneratorTests
{
    [Fact]
    public void Generate_GroupsIdenticalPipesIntoOneRow()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "12"),
        };

        var result = generator.Generate(components, profile);

        var pipeRow = Assert.Single(result.Rows.Where(row => row.RowType == BomRowType.PipeCut));
        Assert.Equal(2m, pipeRow.Quantity);
    }

    [Fact]
    public void Generate_SeparatesPipesWithDifferentLengths()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "24"),
        };

        var result = generator.Generate(components, profile);

        Assert.Equal(2, result.Rows.Count(row => row.RowType == BomRowType.PipeCut));
    }

    [Fact]
    public void Generate_CreatesAccessoryRowsUsingPipeQuantities()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", quantity: 2m, numGaskets: "2", numClamps: "1"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "12", quantity: 2m, numGaskets: "2", numClamps: "1"),
        };

        var result = generator.Generate(components, profile);

        var gasketRow = Assert.Single(result.Rows.Where(row => row.RowType == BomRowType.Accessory && row.Values["Description"] == "Gaskets"));
        var clampRow = Assert.Single(result.Rows.Where(row => row.RowType == BomRowType.Accessory && row.Values["Description"] == "Clamps"));

        Assert.Equal(8m, gasketRow.Quantity);
        Assert.Equal(4m, clampRow.Quantity);
        Assert.All(result.Rows.Where(row => row.RowType == BomRowType.Accessory), row => Assert.Equal(KnownBomSections.PipeAccessories, row.Section));
    }

    [Fact]
    public void Generate_DoesNotExposeIgnoredPropertiesAsColumns()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var result = generator.Generate([TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12")], profile);

        var pipeRow = Assert.Single(result.Rows.Where(row => row.RowType == BomRowType.PipeCut));

        Assert.DoesNotContain(pipeRow.Values.Keys, key => string.Equals(key, KnownPropertyNames.BlueGasket, StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(pipeRow.Values.Keys, key => string.Equals(key, KnownPropertyNames.WhiteGasket, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Generate_ReportsMissingPipeLengthForPipeCandidate()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var component = new ComponentRecord
        {
            ComponentId = "pipe-candidate",
            ComponentName = "pipe-candidate",
            ConfigurationName = "Default",
            Properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
            {
                [KnownPropertyNames.Bom] = TestData.CreateProperty(KnownPropertyNames.Bom, "B-100"),
                [KnownPropertyNames.PipeIdentifier] = TestData.CreateProperty(KnownPropertyNames.PipeIdentifier, "Line A"),
                [KnownPropertyNames.Specification] = TestData.CreateProperty(KnownPropertyNames.Specification, "Spec A"),
                [KnownPropertyNames.NumGaskets] = TestData.CreateProperty(KnownPropertyNames.NumGaskets, "2"),
            },
        };

        var result = generator.Generate([component], profile);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "missing-pipe-length");
    }

    [Fact]
    public void Generate_GroupsVirtualPartsWithoutFilePaths()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new[]
        {
            TestData.CreateVirtualFitting("virt-1", "Clamp", quantity: 1m),
            TestData.CreateVirtualFitting("virt-2", "Clamp", quantity: 1m),
        };

        var result = generator.Generate(components, profile);

        var otherRow = Assert.Single(result.Rows.Where(row => row.RowType == BomRowType.Other));
        Assert.Equal(2m, otherRow.Quantity);
        Assert.Equal("Virtual", otherRow.Values["Component Type"]);
    }
}
