namespace BomCore.Tests;

public sealed class PropertyDiscoveryServiceTests
{
    [Fact]
    public void SuggestDefaultProfile_UsesKnownPipeMappings()
    {
        var service = new PropertyDiscoveryService();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", numGaskets: "2", numClamps: "1"),
        };

        var profile = service.SuggestDefaultProfile(components);

        Assert.Contains(profile.PipeColumns, column => column.DisplayName == "BOM Code");
        Assert.Contains(profile.PipeColumns, column => column.DisplayName == "Pipe Description");
        Assert.Contains(profile.PipeColumns, column => column.DisplayName == "Specification");
        Assert.Contains(profile.PipeColumns, column => column.DisplayName == "Cut Length");
        Assert.Contains(profile.AccessoryRules, rule => rule.DisplayName == "Gaskets");
        Assert.Contains(profile.AccessoryRules, rule => rule.DisplayName == "Clamps");
    }

    [Fact]
    public void DiscoverFromComponents_MarksIgnoredPropertiesInsteadOfRemovingThem()
    {
        var service = new PropertyDiscoveryService();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12"),
        };

        var result = service.DiscoverFromComponents(components);

        Assert.Contains(KnownPropertyNames.BlueGasket, result.DiscoveredProperties);
        Assert.Contains(KnownPropertyNames.WhiteGasket, result.DiscoveredProperties);
        Assert.Contains(KnownPropertyNames.BlueGasket, result.IgnoredProperties);
        Assert.Contains(KnownPropertyNames.WhiteGasket, result.IgnoredProperties);
    }
}
