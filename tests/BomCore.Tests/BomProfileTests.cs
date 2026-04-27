namespace BomCore.Tests;

public sealed class BomProfileTests
{
    [Fact]
    public void Deserialize_LoadsValidProfile()
    {
        var json = """
        {
          "profileName": "AFCA Pipe BOM",
          "version": 2,
          "partClassRules": [
            {
              "className": "Pipe",
              "detectWhenPropertyExists": "PipeLength"
            }
          ],
          "pipeColumns": [
            {
              "sourceProperty": "BOM",
              "displayName": "BOM Code",
              "enabled": true,
              "groupBy": true,
              "order": 1
            },
            {
              "sourceProperty": "Pipe Identifier",
              "displayName": "Pipe Description",
              "enabled": true,
              "groupBy": true,
              "order": 2
            },
            {
              "sourceProperty": "Specification",
              "displayName": "Specification",
              "enabled": true,
              "groupBy": true,
              "order": 3
            },
            {
              "sourceProperty": "Component Type",
              "displayName": "Component Type",
              "enabled": true,
              "groupBy": true,
              "order": 4
            },
            {
              "sourceProperty": "PipeLength",
              "displayName": "Cut Length",
              "enabled": true,
              "groupBy": true,
              "order": 5,
              "unit": "in"
            }
          ],
          "sectionColumnProfiles": [
            {
              "section": "Pipes",
              "columns": [
                {
                  "sourceProperty": "BOM",
                  "displayName": "BOM Code",
                  "enabled": true,
                  "groupBy": true,
                  "order": 1
                },
                {
                  "sourceProperty": "Pipe Identifier",
                  "displayName": "Pipe Description",
                  "enabled": true,
                  "groupBy": true,
                  "order": 2
                },
                {
                  "sourceProperty": "Specification",
                  "displayName": "Specification",
                  "enabled": true,
                  "groupBy": true,
                  "order": 3
                },
                {
                  "sourceProperty": "Component Type",
                  "displayName": "Component Type",
                  "enabled": true,
                  "groupBy": true,
                  "order": 4
                },
                {
                  "sourceProperty": "PipeLength",
                  "displayName": "Cut Length",
                  "enabled": true,
                  "groupBy": true,
                  "order": 5,
                  "unit": "in"
                }
              ]
            },
            {
              "section": "Components",
              "columns": [
                {
                  "sourceProperty": "BOMDesc",
                  "displayName": "BOMDesc",
                  "enabled": true,
                  "groupBy": true,
                  "order": 1
                },
                {
                  "sourceProperty": "Description",
                  "displayName": "Description",
                  "enabled": true,
                  "groupBy": true,
                  "order": 2
                }
              ]
            }
          ],
          "sectionRules": [
            {
              "sourceProperty": "Primary Family",
              "matchValue": "Fittings",
              "section": "Fittings"
            }
          ],
          "accessoryRules": [
            {
              "sourceProperty": "NumGaskets",
              "displayName": "Gaskets",
              "bomSection": "Other Accessories"
            },
            {
              "sourceProperty": "NumClamps",
              "displayName": "Clamps",
              "bomSection": "Other Accessories"
            }
          ],
          "ignoredProperties": [
            "BlueGasket",
            "WhiteGasket",
            "BlueFerrule",
            "WhiteFerrule"
          ]
        }
        """;

        var profile = BomProfileSerializer.Deserialize(json);

        Assert.Equal("AFCA Pipe BOM", profile.ProfileName);
        Assert.Equal(5, profile.GetSectionColumns(KnownBomSections.Pipes).Count);
        Assert.Contains(profile.SectionRules, rule => rule.MatchValue == KnownBomSections.Fittings);
        Assert.Equal(KnownPropertyNames.PipeLength, profile.GetPipeDetectionProperty());
    }

    [Fact]
    public void Deserialize_LoadsRepositoryDefaultProfile()
    {
        var profilePath = TestData.GetRepositoryPath("profiles", "default.pipebom.json");
        var profile = BomProfileSerializer.Deserialize(File.ReadAllText(profilePath));

        Assert.Equal("AFCA Pipe BOM", profile.ProfileName);
        Assert.Contains(profile.AccessoryRules, rule => rule.SourceProperty == KnownPropertyNames.NumGaskets);
        Assert.Contains(profile.SectionColumnProfiles, sectionProfile => sectionProfile.Section == KnownBomSections.Components);
    }

    [Fact]
    public void Validate_ReportsDuplicateDisplayNames()
    {
        var profile = TestData.CreateDefaultProfile() with
        {
            AccessoryRules =
            [
                new AccessoryRule
                {
                    SourceProperty = KnownPropertyNames.NumGaskets,
                    DisplayName = "Pipe Description",
                    BomSection = KnownBomSections.PipeAccessories,
                },
            ],
        };

        var diagnostics = BomProfileSerializer.Validate(profile);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "duplicate-display-name");
    }

    [Fact]
    public void Validate_ReportsMissingRequiredPipeFields()
    {
        var profile = TestData.CreateDefaultProfile() with
        {
            PipeColumns =
            [
                new BomColumnRule
                {
                    SourceProperty = KnownPropertyNames.Bom,
                    DisplayName = "BOM Code",
                    Enabled = true,
                    GroupBy = true,
                    Order = 1,
                },
            ],
            SectionColumnProfiles =
            [
                new BomSectionColumnProfile
                {
                    Section = KnownBomSections.Pipes,
                    Columns =
                    [
                        new BomColumnRule
                        {
                            SourceProperty = KnownPropertyNames.Bom,
                            DisplayName = "BOM Code",
                            Enabled = true,
                            GroupBy = true,
                            Order = 1,
                        },
                    ],
                },
            ],
        };

        var diagnostics = BomProfileSerializer.Validate(profile);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "missing-pipe-field" && diagnostic.PropertyName == KnownPropertyNames.PipeLength);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "missing-pipe-field" && diagnostic.PropertyName == KnownPropertyNames.PipeIdentifier);
    }

    [Fact]
    public void Serialize_RoundTripsSectionColumnsAndRules()
    {
        var profile = TestData.CreateDefaultProfile() with
        {
            SectionColumnProfiles =
            [
                new BomSectionColumnProfile
                {
                    Section = KnownBomSections.Fittings,
                    Columns =
                    [
                        new BomColumnRule
                        {
                            SourceProperty = "CustomCode",
                            DisplayName = "Custom Code",
                            Enabled = true,
                            GroupBy = true,
                            Order = 1,
                            Unit = "ea",
                        },
                    ],
                },
            ],
            SectionRules =
            [
                new BomSectionRule
                {
                    SourceProperty = KnownPropertyNames.PrimaryFamily,
                    MatchValue = "Valve",
                    Section = KnownBomSections.Fittings,
                },
            ],
        };

        var roundTripped = BomProfileSerializer.Deserialize(BomProfileSerializer.Serialize(profile));

        var sectionProfile = Assert.Single(roundTripped.SectionColumnProfiles);
        Assert.Equal(KnownBomSections.Fittings, sectionProfile.Section);
        Assert.Equal("CustomCode", Assert.Single(sectionProfile.Columns).SourceProperty);
        var sectionRule = Assert.Single(roundTripped.SectionRules);
        Assert.Equal("Valve", sectionRule.MatchValue);
        Assert.Equal(KnownBomSections.Fittings, sectionRule.Section);
    }

    [Fact]
    public void Deserialize_MigratesOldPipeColumnsToPipesSection()
    {
        var profile = new BomProfile
        {
            PipeColumns =
            [
                new BomColumnRule
                {
                    SourceProperty = KnownPropertyNames.PipeLength,
                    DisplayName = "Cut Length",
                    Enabled = true,
                    GroupBy = true,
                    Order = 1,
                },
            ],
        };

        var pipeColumn = Assert.Single(profile.GetSectionColumns(KnownBomSections.Pipes));
        Assert.Equal(KnownPropertyNames.PipeLength, pipeColumn.SourceProperty);
    }
}
