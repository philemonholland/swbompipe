namespace BomCore.Tests;

public sealed class ProfileStoreTests
{
    [Fact]
    public void LoadEffectiveProfile_PrefersProjectLocalProfile()
    {
        var root = TestData.CreateTempDirectory();

        try
        {
            var store = new ProfileStore();
            var defaultProfilePath = Path.Combine(root, "defaults", "default.pipebom.json");
            var projectDirectory = Path.Combine(root, "project");
            Directory.CreateDirectory(Path.GetDirectoryName(defaultProfilePath)!);
            Directory.CreateDirectory(projectDirectory);

            store.SaveToPath(TestData.CreateDefaultProfile(), defaultProfilePath);

            var projectProfile = TestData.CreateDefaultProfile() with { ProfileName = "Project Profile" };
            var projectProfilePath = Path.Combine(projectDirectory, "default.pipebom.json");
            store.SaveToPath(projectProfile, projectProfilePath);

            var result = store.LoadEffectiveProfile(
                Path.Combine(projectDirectory, "Test_Project.SLDASM"),
                new ProfileStoreOptions
                {
                    DefaultProfilePath = defaultProfilePath,
                    UserProfileDirectory = Path.Combine(root, "user"),
                    CompanyProfileDirectory = Path.Combine(root, "company"),
                });

            Assert.Equal(projectProfilePath, result.SourcePath);
            Assert.Equal("Project Profile", result.Profile.ProfileName);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void LoadEffectiveProfile_FallsBackToDefaultWhenCandidateIsInvalid()
    {
        var root = TestData.CreateTempDirectory();

        try
        {
            var store = new ProfileStore();
            var defaultProfilePath = Path.Combine(root, "defaults", "default.pipebom.json");
            var userDirectory = Path.Combine(root, "user");
            Directory.CreateDirectory(Path.GetDirectoryName(defaultProfilePath)!);
            Directory.CreateDirectory(userDirectory);

            store.SaveToPath(TestData.CreateDefaultProfile(), defaultProfilePath);
            File.WriteAllText(Path.Combine(userDirectory, "default.pipebom.json"), "{ invalid json");

            var result = store.LoadEffectiveProfile(
                Path.Combine(root, "project", "Test_Project.SLDASM"),
                new ProfileStoreOptions
                {
                    DefaultProfilePath = defaultProfilePath,
                    UserProfileDirectory = userDirectory,
                });

            Assert.Equal(defaultProfilePath, result.SourcePath);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "profile-load-failed");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void LoadDefaultProfile_PrefersUserProfileWhenNoAssemblyPathIsAvailable()
    {
        var root = TestData.CreateTempDirectory();

        try
        {
            var store = new ProfileStore();
            var defaultProfilePath = Path.Combine(root, "defaults", "default.pipebom.json");
            var userDirectory = Path.Combine(root, "user");
            Directory.CreateDirectory(Path.GetDirectoryName(defaultProfilePath)!);
            Directory.CreateDirectory(userDirectory);

            store.SaveToPath(TestData.CreateDefaultProfile(), defaultProfilePath);

            var userProfile = TestData.CreateDefaultProfile() with { ProfileName = "User Default Profile" };
            var userProfilePath = Path.Combine(userDirectory, "default.pipebom.json");
            store.SaveToPath(userProfile, userProfilePath);

            var result = store.LoadDefaultProfile(
                new ProfileStoreOptions
                {
                    DefaultProfilePath = defaultProfilePath,
                    UserProfileDirectory = userDirectory,
                    CompanyProfileDirectory = Path.Combine(root, "company"),
                });

            Assert.Equal(userProfilePath, result.SourcePath);
            Assert.Equal("User Default Profile", result.Profile.ProfileName);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
