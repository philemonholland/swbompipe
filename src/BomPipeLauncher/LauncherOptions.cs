using BomCore;

namespace BomPipeLauncher;

internal sealed record LauncherOptions
{
    public string AssemblyPath { get; init; } = string.Empty;

    public string? ProfilePath { get; init; }

    public string? OutputPath { get; init; }

    public string? DebugReportPath { get; init; }

    public string Format { get; init; } = "csv";

    public bool Visible { get; init; }

    public static LauncherOptions Parse(string[] args)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index];
            if (!argument.StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Unexpected argument '{argument}'.");
            }

            var key = argument[2..];
            if (string.Equals(key, "visible", StringComparison.OrdinalIgnoreCase))
            {
                flags.Add(key);
                continue;
            }

            if (index == args.Length - 1)
            {
                throw new ArgumentException($"Missing value for '{argument}'.");
            }

            values[key] = args[++index];
        }

        if (!values.TryGetValue("assembly", out var assemblyPath) || string.IsNullOrWhiteSpace(assemblyPath))
        {
            throw new ArgumentException("An assembly path is required. Use --assembly <path>.");
        }

        var format = values.TryGetValue("format", out var explicitFormat) ? explicitFormat : "csv";
        if (!string.Equals(format, "csv", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Format must be either 'csv' or 'xlsx'.");
        }

        var debugReportPath = values.TryGetValue("debug-report", out var explicitDebugReportPath)
            ? Path.GetFullPath(explicitDebugReportPath)
            : null;
        if (debugReportPath is not null)
        {
            DebugReportExporter.GetFormatFromPath(debugReportPath);
        }

        return new LauncherOptions
        {
            AssemblyPath = Path.GetFullPath(assemblyPath),
            ProfilePath = values.TryGetValue("profile", out var profilePath) ? Path.GetFullPath(profilePath) : null,
            OutputPath = values.TryGetValue("output", out var outputPath) ? Path.GetFullPath(outputPath) : null,
            DebugReportPath = debugReportPath,
            Format = format.ToLowerInvariant(),
            Visible = flags.Contains("visible"),
        };
    }

    public static string Usage =>
        """
        Usage:
          BomPipeLauncher --assembly <path-to-.SLDASM> [--profile <path-to-json>] [--output <path>] [--format csv|xlsx] [--debug-report <path-to-.txt-or-.json>] [--visible]

        Examples:
          BomPipeLauncher --assembly "D:\AFCA\projects\test_project_2\Test_Project.SLDASM"
          BomPipeLauncher --assembly "D:\AFCA\projects\test_project_2\Test_Project.SLDASM" --format xlsx --output "D:\Exports\Test_Project.xlsx"
          BomPipeLauncher --assembly "D:\AFCA\projects\test_project_2\Test_Project.SLDASM" --debug-report "D:\Exports\Test_Project.debug.json"
        """;
}
