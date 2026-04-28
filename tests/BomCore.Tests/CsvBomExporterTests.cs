namespace BomCore.Tests;

public sealed class CsvBomExporterTests
{
    [Fact]
    public void Export_WritesSectionsHeadersAndQuantities()
    {
        var exporter = new CsvBomExporter();
        var result = new BomResult
        {
            Rows =
            [
                new BomRow
                {
                    Section = KnownBomSections.PipeCutList,
                    RowType = BomRowType.PipeCut,
                    Quantity = 2m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["BOM Code"] = "B-100",
                        ["Pipe Description"] = "Line A",
                    },
                },
                new BomRow
                {
                    Section = KnownBomSections.PipeAccessories,
                    RowType = BomRowType.Accessory,
                    Quantity = 8m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Description"] = "Gaskets",
                    },
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();

        exporter.Export(result, stream);

        var csv = TestData.ReadUtf8(stream);

        Assert.Contains("Pipes", csv);
        Assert.Contains("Other Accessories", csv);
        Assert.Contains("BOM Code,Pipe Description,Quantity", csv);
        Assert.Contains("B-100,Line A,2", csv);
        Assert.Contains("Description,Quantity", csv);
        Assert.Contains("Gaskets,8", csv);
    }

    [Fact]
    public void Export_WritesDynamicSectionsBeforeOtherAndAccessories()
    {
        var exporter = new CsvBomExporter();
        var result = new BomResult
        {
            Rows =
            [
                new BomRow
                {
                    Section = "Valves",
                    RowType = BomRowType.SectionItem,
                    Quantity = 1m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Description"] = "Valve A",
                    },
                },
                new BomRow
                {
                    Section = KnownBomSections.Other,
                    RowType = BomRowType.SectionItem,
                    Quantity = 1m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Description"] = "Other Part",
                    },
                },
                new BomRow
                {
                    Section = KnownBomSections.OtherAccessories,
                    RowType = BomRowType.Accessory,
                    Quantity = 2m,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Description"] = "Gaskets",
                    },
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();

        exporter.Export(result, stream);

        var csv = TestData.ReadUtf8(stream);
        var valvesIndex = csv.IndexOf("Valves", StringComparison.Ordinal);
        var otherIndex = csv.IndexOf("Other", StringComparison.Ordinal);
        var accessoriesIndex = csv.IndexOf("Other Accessories", StringComparison.Ordinal);

        Assert.True(valvesIndex >= 0);
        Assert.True(otherIndex > valvesIndex);
        Assert.True(accessoriesIndex > otherIndex);
    }
}
