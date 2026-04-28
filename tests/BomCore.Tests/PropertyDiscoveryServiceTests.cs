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

        var pipeColumns = profile.GetSectionColumns(KnownBomSections.Pipes);
        Assert.Contains(pipeColumns, column => column.DisplayName == "BOM Code");
        Assert.Contains(pipeColumns, column => column.DisplayName == "Pipe Description");
        Assert.Contains(pipeColumns, column => column.DisplayName == "Specification");
        Assert.Contains(pipeColumns, column => column.DisplayName == "Cut Length");
        Assert.Contains(profile.SectionColumnProfiles, sectionProfile => sectionProfile.Section == KnownBomSections.Tubes);
        Assert.Contains(profile.SectionRules, rule => rule.SourceProperty == KnownPropertyNames.PrimaryFamily && rule.Section == KnownBomSections.Fittings);
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

    [Fact]
    public void DiscoverFromComponents_ReportsCustomPrimaryFamilySections()
    {
        var service = new PropertyDiscoveryService();
        var components = new[]
        {
            TestData.CreateClassifiedComponent("valve-1", "Valves", "VL-100", "Valve A"),
        };

        var result = service.DiscoverFromComponents(components);

        Assert.Contains("Valves", result.DiscoveredSections);
        Assert.DoesNotContain(KnownBomSections.Pipes, result.DiscoveredSections);
        Assert.DoesNotContain(KnownBomSections.Other, result.DiscoveredSections);
    }

    [Fact]
    public void SuggestDefaultProfile_AddsCustomPrimaryFamilySectionProfiles()
    {
        var service = new PropertyDiscoveryService();
        var components = new[]
        {
            TestData.CreateClassifiedComponent("valve-1", "Valves", "VL-100", "Valve A"),
        };

        var profile = service.SuggestDefaultProfile(components);

        Assert.Contains(profile.SectionColumnProfiles, sectionProfile => string.Equals(sectionProfile.Section, "Valves", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(
            ["BOMDesc", "Description", "Component Type"],
            profile.GetSectionColumns("Valves").Select(column => column.DisplayName).ToArray());
    }
}
