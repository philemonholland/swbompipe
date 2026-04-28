using System.Reflection;

namespace SolidWorksBOMAddin;

internal static class BuiltInProfileResolver
{
    private const string DefaultProfileFileName = "default.pipebom.json";
    private const string EmbeddedProfileResourceName = "SolidWorksBOMAddin.default.pipebom.json";

    public static string ResolvePath()
    {
        var candidatePaths = EnumerateCandidatePaths()
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var candidatePath in candidatePaths)
        {
            if (!File.Exists(candidatePath))
            {
                continue;
            }

            BomPipeLog.Info($"Using built-in BOM profile from '{candidatePath}'.");
            return candidatePath;
        }

        return MaterializeEmbeddedProfile(candidatePaths);
    }

    private static IEnumerable<string> EnumerateCandidatePaths()
    {
        var executingAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var appContextDirectory = AppContext.BaseDirectory;

        foreach (var baseDirectory in new[] { executingAssemblyDirectory, appContextDirectory })
        {
            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                continue;
            }

            yield return Path.Combine(baseDirectory, "profiles", DefaultProfileFileName);

            var installRoot = Directory.GetParent(baseDirectory)?.FullName;
            if (string.IsNullOrWhiteSpace(installRoot))
            {
                continue;
            }

            yield return Path.Combine(installRoot, "SolidWorksBOMAddin", "profiles", DefaultProfileFileName);
            yield return Path.Combine(installRoot, "BomPipeLauncher", "profiles", DefaultProfileFileName);
        }
    }

    private static string MaterializeEmbeddedProfile(IReadOnlyList<string> candidatePaths)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var embeddedStream = assembly.GetManifestResourceStream(EmbeddedProfileResourceName)
            ?? FindEmbeddedProfileStream(assembly);
        if (embeddedStream is null)
        {
            throw new InvalidOperationException(
                $"Could not locate the built-in BOM profile. Tried: {string.Join("; ", candidatePaths)}");
        }

        var materializedPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AFCA",
            "BOMPipe",
            "cache",
            "profiles",
            DefaultProfileFileName);
        var directory = Path.GetDirectoryName(materializedPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using (var fileStream = File.Create(materializedPath))
        {
            embeddedStream.CopyTo(fileStream);
        }

        BomPipeLog.Info(
            $"Built-in BOM profile was not found in staged payload paths. Materialized embedded default profile to '{materializedPath}'.");
        return materializedPath;
    }

    private static Stream? FindEmbeddedProfileStream(Assembly assembly)
    {
        var resourceName = assembly.GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(DefaultProfileFileName, StringComparison.OrdinalIgnoreCase));
        return resourceName is null ? null : assembly.GetManifestResourceStream(resourceName);
    }
}
