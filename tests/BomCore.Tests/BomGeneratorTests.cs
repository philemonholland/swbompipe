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
    public void Generate_MergesIdenticalAccessoryRowsAcrossSeparatePipeGroups()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", numGaskets: "1", numClamps: "1"),
            TestData.CreatePipe("pipe-2", "B-100", "Line A", "Spec A", "24", numGaskets: "1", numClamps: "1"),
        };

        var result = generator.Generate(components, profile);

        var accessoryRows = result.Rows.Where(row => row.RowType == BomRowType.Accessory).ToList();
        var gasketRow = Assert.Single(accessoryRows.Where(row => row.Values["Description"] == "Gaskets"));
        var clampRow = Assert.Single(accessoryRows.Where(row => row.Values["Description"] == "Clamps"));

        Assert.Equal(2m, gasketRow.Quantity);
        Assert.Equal(2m, clampRow.Quantity);
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
    public void Generate_MissingPrimaryFamilyRoutesPipeLikePartToOtherWithoutPipeLengthDiagnostic()
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

        Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Other));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "missing-pipe-length");
    }

    [Fact]
    public void Generate_PipeFamilyMissingPipeLengthReportsConfiguredColumnDiagnostic()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var component = new ComponentRecord
        {
            ComponentId = "pipe-no-length",
            ComponentName = "pipe-no-length",
            ConfigurationName = "Default",
            Properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
            {
                [KnownPropertyNames.PrimaryFamily] = TestData.CreateProperty(KnownPropertyNames.PrimaryFamily, KnownBomSections.Pipes),
                [KnownPropertyNames.Bom] = TestData.CreateProperty(KnownPropertyNames.Bom, "B-100"),
                [KnownPropertyNames.PipeIdentifier] = TestData.CreateProperty(KnownPropertyNames.PipeIdentifier, "Line A"),
                [KnownPropertyNames.ComponentType] = TestData.CreateProperty(KnownPropertyNames.ComponentType, "Pipe"),
                [KnownPropertyNames.Specification] = TestData.CreateProperty(KnownPropertyNames.Specification, "Spec A"),
            },
        };

        var result = generator.Generate([component], profile);

        Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Pipes));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "missing-configured-property"
            && diagnostic.PropertyName == KnownPropertyNames.PipeLength);
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

        var otherRow = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Fittings));
        Assert.Equal(2m, otherRow.Quantity);
        Assert.Equal(KnownBomSections.Fittings, otherRow.Section);
        Assert.Equal("Clamp", otherRow.Values[KnownPropertyNames.BomDesc]);
    }

    [Fact]
    public void Generate_DoesNotExposeComponentFileOrConfigurationFallbacks()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var component = TestData.CreateClassifiedComponent("fit-1", KnownBomSections.Fittings, "FT-100", "Elbow");

        var result = generator.Generate([component], profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Fittings));
        Assert.DoesNotContain(row.Values.Keys, key => string.Equals(key, "Component", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(row.Values.Keys, key => string.Equals(key, "File Path", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(row.Values.Keys, key => string.Equals(key, "Configuration", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(row.Values.Values, value => value.Contains("fit-1", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(row.Values.Values, value => value.Contains(".sldprt", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Generate_TubeLengthRoutesToTubes()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var tube = TestData.CreateClassifiedComponent(
            "tube-1",
            KnownBomSections.Tubes,
            "TB-100",
            "Tube A",
            extraProperties: new Dictionary<string, string>
            {
                [KnownPropertyNames.TubeLength] = "18",
            });

        var result = generator.Generate([tube], profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Tubes));
        Assert.Equal("18", row.Values["Tube Length"]);
    }

    [Fact]
    public void Generate_WireLengthRoutesToWires()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var wire = TestData.CreateClassifiedComponent(
            "wire-1",
            KnownBomSections.Wires,
            "WR-100",
            "Wire A",
            extraProperties: new Dictionary<string, string>
            {
                [KnownPropertyNames.WireLength] = "36",
            });

        var result = generator.Generate([wire], profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Wires));
        Assert.Equal("36", row.Values["Wire Length"]);
    }

    [Fact]
    public void Generate_PrimaryFamilyFittingsRoutesToFittings()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var fitting = TestData.CreateClassifiedComponent("fit-1", KnownBomSections.Fittings, "FT-100", "Elbow");

        var result = generator.Generate([fitting], profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Fittings));
        Assert.Equal("FT-100", row.Values[KnownPropertyNames.BomDesc]);
        Assert.Equal("Elbow", row.Values[KnownPropertyNames.Description]);
        Assert.Equal(1m, row.Quantity);
    }

    [Fact]
    public void Generate_PrimaryFamilyAliasRoutesToRequestedSection()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var fitting = TestData.CreateClassifiedComponent("fit-1", "Fitting", "FT-100", "Elbow");

        var result = generator.Generate([fitting], profile);

        Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Fittings));
    }

    [Fact]
    public void Generate_CustomPrimaryFamilyBecomesOwnSection()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var component = TestData.CreateClassifiedComponent("valve-1", "Valves", "VL-100", "Valve Part");

        var result = generator.Generate([component], profile);

        var row = Assert.Single(result.Rows.Where(row => string.Equals(row.Section, "Valves", StringComparison.OrdinalIgnoreCase)));
        Assert.Equal("VL-100", row.Values[KnownPropertyNames.BomDesc]);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "unmapped-primary-family");
    }

    [Fact]
    public void Generate_DynamicSectionsOrderBeforeOtherAndAccessories()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile();
        var components = new ComponentRecord[]
        {
            TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", numGaskets: "1"),
            TestData.CreateClassifiedComponent("valve-1", "Valves", "VL-100", "Valve Part"),
            new()
            {
                ComponentId = "other-1",
                ComponentName = "other-1",
                ConfigurationName = "Default",
                Properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
                {
                    [KnownPropertyNames.BomDesc] = TestData.CreateProperty(KnownPropertyNames.BomDesc, "Other Part"),
                    [KnownPropertyNames.Description] = TestData.CreateProperty(KnownPropertyNames.Description, "Other Part"),
                },
            },
        };

        var result = generator.Generate(components, profile);

        Assert.Equal(
            [KnownBomSections.Pipes, "Valves", KnownBomSections.Other, KnownBomSections.OtherAccessories],
            result.Rows.Select(row => row.Section).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
    }

    [Fact]
    public void Generate_MissingConfiguredPropertyOutputsBlankAndDiagnostic()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile() with
        {
            SectionColumnProfiles =
            [
                new BomSectionColumnProfile
                {
                    Section = KnownBomSections.Components,
                    Columns =
                    [
                        new BomColumnRule
                        {
                            SourceProperty = "MissingManualProperty",
                            DisplayName = "Manual Header",
                            Enabled = true,
                            GroupBy = true,
                            Order = 1,
                        },
                    ],
                },
            ],
        };
        var component = TestData.CreateClassifiedComponent("comp-1", KnownBomSections.Components, "CP-100", "Component A");

        var result = generator.Generate([component], profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Components));
        Assert.Equal(string.Empty, row.Values["Manual Header"]);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "missing-configured-property"
            && diagnostic.PropertyName == "MissingManualProperty");
    }

    [Fact]
    public void Generate_GroupsByConfiguredColumnsPerSection()
    {
        var generator = new BomGenerator();
        var profile = TestData.CreateDefaultProfile() with
        {
            SectionColumnProfiles =
            [
                new BomSectionColumnProfile
                {
                    Section = KnownBomSections.Components,
                    Columns =
                    [
                        new BomColumnRule
                        {
                            SourceProperty = KnownPropertyNames.BomDesc,
                            DisplayName = KnownPropertyNames.BomDesc,
                            Enabled = true,
                            GroupBy = true,
                            Order = 1,
                        },
                        new BomColumnRule
                        {
                            SourceProperty = KnownPropertyNames.Description,
                            DisplayName = KnownPropertyNames.Description,
                            Enabled = true,
                            GroupBy = false,
                            Order = 2,
                        },
                    ],
                },
            ],
        };
        var components = new[]
        {
            TestData.CreateClassifiedComponent("comp-1", KnownBomSections.Components, "CP-100", "First", quantity: 2m),
            TestData.CreateClassifiedComponent("comp-2", KnownBomSections.Components, "CP-100", "Second", quantity: 3m),
        };

        var result = generator.Generate(components, profile);

        var row = Assert.Single(result.Rows.Where(row => row.Section == KnownBomSections.Components));
        Assert.Equal(5m, row.Quantity);
        Assert.Equal("CP-100", row.Values[KnownPropertyNames.BomDesc]);
    }
}
