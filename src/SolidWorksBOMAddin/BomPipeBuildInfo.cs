using System.Reflection;
using System.Text.Json;

namespace SolidWorksBOMAddin;

internal sealed record BomPipeBuildInfo(
    string SourceRevision,
    string InstalledAtUtc,
    string Configuration)
{
    public string DisplayText
    {
        get
        {
            var revision = string.IsNullOrWhiteSpace(SourceRevision) ? "unknown" : SourceRevision;
            var installed = string.IsNullOrWhiteSpace(InstalledAtUtc) ? "unknown install time" : InstalledAtUtc;
            var configuration = string.IsNullOrWhiteSpace(Configuration) ? "unknown" : Configuration;
            return $"Build: {revision} | {configuration} | Installed: {installed}";
        }
    }

    public bool IsVerified => !SourceRevision.Contains("reinstall BOMPipe", StringComparison.OrdinalIgnoreCase);

    public static BomPipeBuildInfo Load()
    {
        var assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var manifestPath = string.IsNullOrWhiteSpace(assemblyDirectory)
            ? null
            : Path.Combine(assemblyDirectory, "bompipe-build-info.json");

        if (string.IsNullOrWhiteSpace(manifestPath) || !File.Exists(manifestPath))
        {
            return new BomPipeBuildInfo(
                "unknown - reinstall BOMPipe",
                "manifest missing",
                "unknown");
        }

        try
        {
            using var stream = File.OpenRead(manifestPath);
            using var document = JsonDocument.Parse(stream);
            var root = document.RootElement;
            return new BomPipeBuildInfo(
                GetString(root, "source_revision"),
                GetString(root, "installed_at_utc"),
                GetString(root, "configuration"));
        }
        catch (IOException)
        {
            return new BomPipeBuildInfo("unreadable - reinstall BOMPipe", manifestPath, "unknown");
        }
        catch (JsonException)
        {
            return new BomPipeBuildInfo("invalid manifest - reinstall BOMPipe", manifestPath, "unknown");
        }
    }

    private static string GetString(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;
    }
}
