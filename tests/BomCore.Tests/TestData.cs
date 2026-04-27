using System.Text;

namespace BomCore.Tests;

internal static class TestData
{
    public static BomProfile CreateDefaultProfile()
    {
        return new BomProfile
        {
            ProfileName = "AFCA Pipe BOM",
            Version = 2,
            PartClassRules =
            [
                new PartClassRule
                {
                    ClassName = "Pipe",
                    DetectWhenPropertyExists = KnownPropertyNames.PipeLength,
                },
            ],
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
                new BomColumnRule
                {
                    SourceProperty = KnownPropertyNames.PipeIdentifier,
                    DisplayName = "Pipe Description",
                    Enabled = true,
                    GroupBy = true,
                    Order = 2,
                },
                new BomColumnRule
                {
                    SourceProperty = KnownPropertyNames.Specification,
                    DisplayName = "Specification",
                    Enabled = true,
                    GroupBy = true,
                    Order = 3,
                },
                new BomColumnRule
                {
                    SourceProperty = KnownPropertyNames.PipeLength,
                    DisplayName = "Cut Length",
                    Enabled = true,
                    GroupBy = true,
                    Order = 4,
                    Unit = "in",
                },
            ],
            SectionColumnProfiles = KnownBomSections.ConfigurableSections
                .Select(section => new BomSectionColumnProfile
                {
                    Section = section,
                    Columns = section == KnownBomSections.Pipes
                        ? KnownBomColumnProfiles.CreateDefaultSectionColumns(KnownBomSections.Pipes)
                        : KnownBomColumnProfiles.CreateDefaultSectionColumns(section),
                })
                .ToList(),
            SectionRules = KnownBomSections.ClassMappedSections
                .Select(section => new BomSectionRule
                {
                    SourceProperty = KnownPropertyNames.PrimaryFamily,
                    MatchValue = section,
                    Section = section,
                })
                .ToList(),
            AccessoryRules =
            [
                new AccessoryRule
                {
                    SourceProperty = KnownPropertyNames.NumGaskets,
                    DisplayName = "Gaskets",
                    BomSection = KnownBomSections.PipeAccessories,
                },
                new AccessoryRule
                {
                    SourceProperty = KnownPropertyNames.NumClamps,
                    DisplayName = "Clamps",
                    BomSection = KnownBomSections.PipeAccessories,
                },
            ],
            IgnoredProperties = KnownPropertyNames.DefaultIgnoredProperties.ToList(),
        };
    }

    public static ComponentRecord CreatePipe(
        string componentId,
        string bom,
        string pipeIdentifier,
        string specification,
        string pipeLength,
        decimal quantity = 1m,
        string? numGaskets = null,
        string? numClamps = null)
    {
        var properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
        {
            [KnownPropertyNames.Bom] = CreateProperty(KnownPropertyNames.Bom, bom),
            [KnownPropertyNames.PrimaryFamily] = CreateProperty(KnownPropertyNames.PrimaryFamily, KnownBomSections.Pipes),
            [KnownPropertyNames.ComponentType] = CreateProperty(KnownPropertyNames.ComponentType, "Pipe"),
            [KnownPropertyNames.PipeIdentifier] = CreateProperty(KnownPropertyNames.PipeIdentifier, pipeIdentifier),
            [KnownPropertyNames.Specification] = CreateProperty(KnownPropertyNames.Specification, specification),
            [KnownPropertyNames.PipeLength] = CreateProperty(KnownPropertyNames.PipeLength, pipeLength),
            [KnownPropertyNames.BlueGasket] = CreateProperty(KnownPropertyNames.BlueGasket, "ignored"),
            [KnownPropertyNames.WhiteGasket] = CreateProperty(KnownPropertyNames.WhiteGasket, "ignored"),
        };

        if (numGaskets is not null)
        {
            properties[KnownPropertyNames.NumGaskets] = CreateProperty(KnownPropertyNames.NumGaskets, numGaskets);
        }

        if (numClamps is not null)
        {
            properties[KnownPropertyNames.NumClamps] = CreateProperty(KnownPropertyNames.NumClamps, numClamps);
        }

        return new ComponentRecord
        {
            ComponentId = componentId,
            FilePath = $@"C:\assemblies\{componentId}.sldprt",
            ConfigurationName = "Default",
            ComponentName = componentId,
            Quantity = quantity,
            Properties = properties,
        };
    }

    public static ComponentRecord CreateVirtualFitting(string componentId, string componentName, decimal quantity = 1m)
    {
        return new ComponentRecord
        {
            ComponentId = componentId,
            ParentComponentId = "subassembly-1",
            ConfigurationName = "Default",
            ComponentName = componentName,
            Quantity = quantity,
            IsVirtual = true,
            Properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
            {
                [KnownPropertyNames.PrimaryFamily] = CreateProperty(KnownPropertyNames.PrimaryFamily, KnownBomSections.Fittings),
                [KnownPropertyNames.ComponentType] = CreateProperty(KnownPropertyNames.ComponentType, "Virtual"),
                [KnownPropertyNames.BomDesc] = CreateProperty(KnownPropertyNames.BomDesc, componentName),
                [KnownPropertyNames.Description] = CreateProperty(KnownPropertyNames.Description, componentName),
            },
        };
    }

    public static ComponentRecord CreateClassifiedComponent(
        string componentId,
        string section,
        string bomDesc,
        string description,
        decimal quantity = 1m,
        IReadOnlyDictionary<string, string>? extraProperties = null)
    {
        var properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
        {
            [KnownPropertyNames.Class] = CreateProperty(KnownPropertyNames.Class, section),
            [KnownPropertyNames.PrimaryFamily] = CreateProperty(KnownPropertyNames.PrimaryFamily, section),
            [KnownPropertyNames.ComponentType] = CreateProperty(KnownPropertyNames.ComponentType, section),
            [KnownPropertyNames.BomDesc] = CreateProperty(KnownPropertyNames.BomDesc, bomDesc),
            [KnownPropertyNames.Description] = CreateProperty(KnownPropertyNames.Description, description),
        };

        if (extraProperties is not null)
        {
            foreach (var pair in extraProperties)
            {
                properties[pair.Key] = CreateProperty(pair.Key, pair.Value);
            }
        }

        return new ComponentRecord
        {
            ComponentId = componentId,
            FilePath = $@"C:\assemblies\{componentId}.sldprt",
            ConfigurationName = "Default",
            ComponentName = componentId,
            Quantity = quantity,
            Properties = properties,
        };
    }

    public static PropertyValue CreateProperty(string name, string value)
    {
        return new PropertyValue
        {
            Name = name,
            RawValue = value,
            EvaluatedValue = value,
            Scope = PropertyScope.Configuration,
            Source = "Test",
        };
    }

    public static string CreateTempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bompipe-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    public static MemoryStream CreateWritableStream()
    {
        return new MemoryStream();
    }

    public static string ReadUtf8(MemoryStream stream)
    {
        stream.Position = 0;
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    public static string GetRepositoryPath(params string[] relativeSegments)
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null && !File.Exists(Path.Combine(current.FullName, "SolidWorksBOMAddin.sln")))
        {
            current = current.Parent;
        }

        if (current is null)
        {
            throw new InvalidOperationException("Could not locate the repository root from the test output directory.");
        }

        return Path.Combine([current.FullName, .. relativeSegments]);
    }
}
