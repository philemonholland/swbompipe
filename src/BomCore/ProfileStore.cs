namespace BomCore;

public sealed class ProfileStore
{
    public BomProfile LoadFromPath(string path)
    {
        return BomProfileSerializer.Deserialize(File.ReadAllText(path));
    }

    public async Task<BomProfile> LoadFromPathAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = File.OpenRead(path);
        return await BomProfileSerializer.DeserializeAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    public void SaveToPath(BomProfile profile, string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(path, BomProfileSerializer.Serialize(profile));
    }

    public ProfileLoadResult LoadEffectiveProfile(string assemblyPath, ProfileStoreOptions options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(assemblyPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(options.DefaultProfilePath);

        var diagnostics = new List<BomDiagnostic>();
        var assemblyDirectory = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
        var candidatePaths = new[]
        {
            Path.Combine(assemblyDirectory, options.DefaultProfileFileName),
            TryBuildCandidatePath(options.UserProfileDirectory, options.DefaultProfileFileName),
            TryBuildCandidatePath(options.CompanyProfileDirectory, options.DefaultProfileFileName),
        };

        foreach (var candidatePath in candidatePaths.Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path)))
        {
            try
            {
                var profile = LoadFromPath(candidatePath!);
                var profileDiagnostics = BomProfileSerializer.Validate(profile);
                if (profileDiagnostics.Any(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error))
                {
                    diagnostics.Add(new BomDiagnostic
                    {
                        Severity = DiagnosticSeverity.Warning,
                        Code = "invalid-profile-fallback",
                        Message = $"Profile '{candidatePath}' is invalid. Falling back to the built-in default profile.",
                    });

                    diagnostics.AddRange(profileDiagnostics);
                    return LoadDefaultProfile(options.DefaultProfilePath, diagnostics);
                }

                diagnostics.AddRange(profileDiagnostics);
                return new ProfileLoadResult
                {
                    Profile = profile,
                    SourcePath = candidatePath,
                    Diagnostics = diagnostics,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Warning,
                    Code = "profile-load-failed",
                    Message = $"Could not load profile '{candidatePath}'. Falling back to the built-in default profile.",
                });

                return LoadDefaultProfile(options.DefaultProfilePath, diagnostics);
            }
        }

        return LoadDefaultProfile(options.DefaultProfilePath, diagnostics);
    }

    private static ProfileLoadResult LoadDefaultProfile(string defaultProfilePath, IReadOnlyList<BomDiagnostic> existingDiagnostics)
    {
        var diagnostics = existingDiagnostics.ToList();
        var profile = BomProfileSerializer.Deserialize(File.ReadAllText(defaultProfilePath));
        diagnostics.AddRange(BomProfileSerializer.Validate(profile));

        return new ProfileLoadResult
        {
            Profile = profile,
            SourcePath = defaultProfilePath,
            Diagnostics = diagnostics,
        };
    }

    private static string? TryBuildCandidatePath(string? directory, string fileName)
    {
        return string.IsNullOrWhiteSpace(directory)
            ? null
            : Path.Combine(directory, fileName);
    }
}
