namespace BomCore.Tests;

public sealed class BomProfileTests
{
    [Fact]
    public void Deserialize_LoadsValidProfile()
    {
        var json = """
        {
          "profileName": "AFCA Pipe BOM",
          "version": 1,
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
              "sourceProperty": "PipeLength",
              "displayName": "Cut Length",
              "enabled": true,
              "groupBy": true,
              "order": 4,
              "unit": "in"
            }
          ],
          "accessoryRules": [
            {
              "sourceProperty": "NumGaskets",
              "displayName": "Gaskets",
              "bomSection": "Pipe Accessories"
            },
            {
              "sourceProperty": "NumClamps",
              "displayName": "Clamps",
              "bomSection": "Pipe Accessories"
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
        Assert.Equal(4, profile.PipeColumns.Count);
        Assert.Equal(KnownPropertyNames.PipeLength, profile.GetPipeDetectionProperty());
    }

    [Fact]
    public void Deserialize_LoadsRepositoryDefaultProfile()
    {
        var profilePath = TestData.GetRepositoryPath("profiles", "default.pipebom.json");
        var profile = BomProfileSerializer.Deserialize(File.ReadAllText(profilePath));

        Assert.Equal("AFCA Pipe BOM", profile.ProfileName);
        Assert.Contains(profile.AccessoryRules, rule => rule.SourceProperty == KnownPropertyNames.NumGaskets);
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
        };

        var diagnostics = BomProfileSerializer.Validate(profile);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "missing-pipe-field" && diagnostic.PropertyName == KnownPropertyNames.PipeLength);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "missing-pipe-field" && diagnostic.PropertyName == KnownPropertyNames.PipeIdentifier);
    }
}
