namespace BomCore.Tests;

public sealed class BomExportFormatTests
{
    [Theory]
    [InlineData(BomExportFormats.Csv, ".bom.csv", ".csv", "CSV BOM")]
    [InlineData(BomExportFormats.Xlsx, ".bom.xlsx", ".xlsx", "Excel BOM")]
    [InlineData(BomExportFormats.BomDbJson, ".bomdb.json", ".json", "BOMDB import JSON")]
    public void FormatMetadata_IsStable(string format, string expectedSuffix, string expectedExtension, string expectedDisplayName)
    {
        Assert.Equal(format, BomExportFormats.Normalize(format));
        Assert.Equal(expectedSuffix, BomExportFormats.GetDefaultFileSuffix(format));
        Assert.Equal(expectedExtension, BomExportFormats.GetDefaultExtension(format));
        Assert.Equal(expectedDisplayName, BomExportFormats.GetDisplayName(format));
    }

    [Theory]
    [InlineData(@"C:\exports\Assembly.bom.csv", BomExportFormats.Csv)]
    [InlineData(@"C:\exports\Assembly.xlsx", BomExportFormats.Xlsx)]
    [InlineData(@"C:\exports\Assembly.bomdb.json", BomExportFormats.BomDbJson)]
    [InlineData(@"C:\exports\Assembly.json", BomExportFormats.BomDbJson)]
    public void TryGetFormatFromPath_RecognizesSupportedExportPaths(string path, string expectedFormat)
    {
        var matched = BomExportFormats.TryGetFormatFromPath(path, out var actualFormat);

        Assert.True(matched);
        Assert.Equal(expectedFormat, actualFormat);
    }

    [Fact]
    public void Normalize_RejectsUnsupportedFormats()
    {
        var exception = Assert.Throws<ArgumentException>(() => BomExportFormats.Normalize("json"));

        Assert.Contains("csv", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("xlsx", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bomdb-json", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
