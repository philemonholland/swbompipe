using System.Text.Json;

namespace BomCore.Tests;

public sealed class DebugReportTests
{
    [Fact]
    public void Create_AggregatesCountsPropertiesAndDiagnostics()
    {
        var service = new DebugReportService();
        BomDiagnostic[] diagnostics =
        [
            new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Warning,
                Code = "component-suppressed",
                Message = "Suppressed component skipped.",
                ComponentId = "suppressed-1",
            },
        ];

        var report = service.Create(
            new DebugReportInput
            {
                AssemblyPath = @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
                ProfilePath = @"C:\codebase\bompipe\profiles\default.pipebom.json",
                Components =
                [
                    TestData.CreatePipe("pipe-1", "B-100", "Line A", "Spec A", "12", numGaskets: "2"),
                    new ComponentRecord
                    {
                        ComponentId = "fitting-1",
                        ComponentName = "Valve",
                        ConfigurationName = "Default",
                        Properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["ValveType"] = TestData.CreateProperty("ValveType", "Gate"),
                        },
                    },
                ],
                ComponentsScanned = 4,
                ComponentsSkipped = 1,
                Rows =
                [
                    new BomRow { RowType = BomRowType.PipeCut, Quantity = 1m },
                    new BomRow { RowType = BomRowType.Other, Quantity = 1m },
                ],
                Diagnostics = diagnostics,
            });

        Assert.Equal(4, report.ComponentsScanned);
        Assert.Equal(1, report.ComponentsSkipped);
        Assert.Equal(2, report.GeneratedBomRowCount);
        Assert.Contains(KnownPropertyNames.PipeLength, report.DiscoveredProperties);
        Assert.Contains(KnownPropertyNames.NumGaskets, report.DiscoveredProperties);
        Assert.Contains("ValveType", report.DiscoveredProperties);
        Assert.Same(diagnostics, report.Diagnostics);
    }

    [Fact]
    public void Export_WritesReadableTextReport()
    {
        var exporter = new DebugReportExporter();
        var report = new DebugReport
        {
            AssemblyPath = @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
            ProfilePath = @"C:\codebase\bompipe\profiles\default.pipebom.json",
            ComponentsScanned = 7,
            ComponentsSkipped = 2,
            DiscoveredProperties = ["BOM", "PipeLength"],
            GeneratedBomRowCount = 3,
            Diagnostics =
            [
                new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Warning,
                    Code = "component-unresolved",
                    Message = "Component 'Pipe-2' was skipped because its model could not be resolved.",
                    ComponentId = "Pipe-2",
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();

        exporter.Export(report, stream, DebugReportFormat.Text);

        var text = TestData.ReadUtf8(stream);

        Assert.Contains("AFCA Piping BOM Debug Report", text);
        Assert.Contains("Assembly Path: D:\\AFCA\\projects\\test_project_2\\Test_Project.SLDASM", text);
        Assert.Contains("Components Scanned: 7", text);
        Assert.Contains("Components Skipped: 2", text);
        Assert.Contains("Generated BOM Rows: 3", text);
        Assert.Contains("Unique Discovered Properties (2):", text);
        Assert.Contains("  - BOM", text);
        Assert.Contains("[Warning] component-unresolved", text);
    }

    [Fact]
    public void Export_WritesJsonReport()
    {
        var exporter = new DebugReportExporter();
        var report = new DebugReport
        {
            AssemblyPath = @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
            ProfilePath = @"C:\codebase\bompipe\profiles\default.pipebom.json",
            ComponentsScanned = 5,
            ComponentsSkipped = 1,
            DiscoveredProperties = ["BOM", "PipeLength"],
            GeneratedBomRowCount = 2,
            Diagnostics =
            [
                new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Info,
                    Code = "component-suppressed",
                    Message = "Suppressed component skipped.",
                },
            ],
        };

        using var stream = TestData.CreateWritableStream();

        exporter.Export(report, stream, DebugReportFormat.Json);

        stream.Position = 0;
        using var document = JsonDocument.Parse(stream);

        Assert.Equal(
            @"D:\AFCA\projects\test_project_2\Test_Project.SLDASM",
            document.RootElement.GetProperty("assemblyPath").GetString());
        Assert.Equal(5, document.RootElement.GetProperty("componentsScanned").GetInt32());
        Assert.Equal(1, document.RootElement.GetProperty("componentsSkipped").GetInt32());
        Assert.Equal(2, document.RootElement.GetProperty("generatedBomRowCount").GetInt32());
        Assert.Equal("component-suppressed", document.RootElement.GetProperty("diagnostics")[0].GetProperty("code").GetString());
    }
}
