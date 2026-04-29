using BomCore;

namespace BomPipeLauncher.Tests;

public sealed class LauncherOptionsTests
{
    [Fact]
    public void Parse_AcceptsBomDbSidecarAlongsideHumanReadablePrimaryOutput()
    {
        var options = LauncherOptions.Parse(
        [
            "--assembly", @"C:\assemblies\TestAssembly.SLDASM",
            "--output", @"C:\exports\TestAssembly.xlsx",
            "--bomdb-output", @"C:\exports\TestAssembly.bomdb.json",
        ]);

        Assert.Equal(Path.GetFullPath(@"C:\assemblies\TestAssembly.SLDASM"), options.AssemblyPath);
        Assert.Equal(BomExportFormats.Xlsx, options.Format);
        Assert.Equal(Path.GetFullPath(@"C:\exports\TestAssembly.xlsx"), options.OutputPath);
        Assert.Equal(Path.GetFullPath(@"C:\exports\TestAssembly.bomdb.json"), options.BomDbOutputPath);
    }

    [Theory]
    [InlineData(@"C:\exports\TestAssembly.txt")]
    [InlineData(@"C:\exports\TestAssembly.csv")]
    [InlineData(@"C:\exports\TestAssembly.xlsx")]
    public void Parse_RejectsNonJsonBomDbOutputPaths(string bomDbOutputPath)
    {
        var exception = Assert.Throws<ArgumentException>(
            () => LauncherOptions.Parse(
            [
                "--assembly", @"C:\assemblies\TestAssembly.SLDASM",
                "--format", "xlsx",
                "--bomdb-output", bomDbOutputPath,
            ]));

        Assert.Contains("bomdb-output", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".json", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Parse_RejectsBomDbSidecarWhenPrimaryOutputIsAlreadyBomDbJson()
    {
        var exception = Assert.Throws<ArgumentException>(
            () => LauncherOptions.Parse(
            [
                "--assembly", @"C:\assemblies\TestAssembly.SLDASM",
                "--format", BomExportFormats.BomDbJson,
                "--output", @"C:\exports\TestAssembly.bomdb.json",
                "--bomdb-output", @"C:\exports\TestAssembly.sidecar.bomdb.json",
            ]));

        Assert.Contains("bomdb-json", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bomdb-output", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
