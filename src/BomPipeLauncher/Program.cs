using System.Runtime.InteropServices;
using BomCore;
using BomPipeLauncher;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksBOMAddin;

return Run(args);

[STAThread]
static int Run(string[] args)
{
    LauncherOptions options;
    try
    {
        options = LauncherOptions.Parse(args);
    }
    catch (ArgumentException ex)
    {
        Console.Error.WriteLine(ex.Message);
        Console.Error.WriteLine();
        Console.Error.WriteLine(LauncherOptions.Usage);
        return 1;
    }

    if (!File.Exists(options.AssemblyPath))
    {
        Console.Error.WriteLine($"Assembly not found: {options.AssemblyPath}");
        return 1;
    }

    if (options.ProfilePath is not null && !File.Exists(options.ProfilePath))
    {
        Console.Error.WriteLine($"Profile not found: {options.ProfilePath}");
        return 1;
    }

    ISldWorks? application = null;
    try
    {
        application = CreateSolidWorksApplication(options.Visible);
        var document = OpenAssembly(application, options.AssemblyPath, options.Visible, out var warnings);
        if (document is null)
        {
            Console.Error.WriteLine("SolidWorks did not return an assembly document.");
            return 1;
        }

        if (warnings != 0)
        {
            Console.WriteLine($"SolidWorks opened the assembly with warnings: {(swFileLoadWarning_e)warnings}");
        }

        var reader = new SolidWorksAssemblyReader(
            application,
            new SolidWorksReaderOptions
            {
                ResolveLightweightComponents = true,
            });

        var store = new ProfileStore();
        var profileLoadResult = LoadProfile(options, store);
        var generator = new BomGenerator();
        var assemblyReadResult = reader.ReadActiveAssembly();
        var bomResult = generator.Generate(assemblyReadResult.Components, profileLoadResult.Profile);
        var combinedDiagnostics = assemblyReadResult.Diagnostics
            .Concat(profileLoadResult.Diagnostics)
            .Concat(bomResult.Diagnostics)
            .ToList();
        var outputPath = ResolveOutputPath(options);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using var outputStream = File.Create(outputPath);
        CreateExporter(options.Format).Export(
            new BomResult
            {
                Rows = bomResult.Rows,
                Diagnostics = combinedDiagnostics,
            },
            outputStream);

        ExportDebugReportIfRequested(options, profileLoadResult, assemblyReadResult, bomResult, combinedDiagnostics);

        Console.WriteLine($"Exported {bomResult.Rows.Count} BOM row(s) to {outputPath}");
        foreach (var diagnostic in combinedDiagnostics)
        {
            Console.WriteLine($"[{diagnostic.Severity}] {diagnostic.Code}: {diagnostic.Message}");
        }

        return 0;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine(ex.Message);
        return 1;
    }
    finally
    {
        if (application is not null)
        {
            try
            {
                application.ExitApp();
            }
            catch
            {
                // Best effort cleanup for externally launched background automation.
            }

            Marshal.FinalReleaseComObject(application);
        }
    }
}

static ISldWorks CreateSolidWorksApplication(bool visible)
{
    var progIdType = Type.GetTypeFromProgID("SldWorks.Application")
        ?? throw new InvalidOperationException("SolidWorks is not installed or its COM registration is unavailable.");
    var application = Activator.CreateInstance(progIdType) as ISldWorks
        ?? throw new InvalidOperationException("Could not create the SolidWorks COM application.");
    application.Visible = visible;
    application.DocumentVisible(visible, (int)swDocumentTypes_e.swDocASSEMBLY);
    return application;
}

static IModelDoc2? OpenAssembly(ISldWorks application, string assemblyPath, bool visible, out int warnings)
{
    application.Visible = visible;
    var errors = 0;
    warnings = 0;
    var document = application.OpenDoc6(
        assemblyPath,
        (int)swDocumentTypes_e.swDocASSEMBLY,
        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
        string.Empty,
        ref errors,
        ref warnings) as IModelDoc2;

    if (document is null)
    {
        throw new InvalidOperationException($"SolidWorks could not open '{assemblyPath}'. Error: {(swFileLoadError_e)errors}.");
    }

    var activationErrors = 0;
    var activatedDocument = application.ActivateDoc3(
        document.GetTitle(),
        UseUserPreferences: false,
        (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
        ref activationErrors) as IModelDoc2;

    if (activatedDocument is not null)
    {
        document = activatedDocument;
    }

    return document;
}

static ProfileLoadResult LoadProfile(LauncherOptions options, ProfileStore store)
{
    if (options.ProfilePath is not null)
    {
        var profile = store.LoadFromPath(options.ProfilePath);
        return new ProfileLoadResult
        {
            Profile = profile,
            SourcePath = options.ProfilePath,
            Diagnostics = BomProfileSerializer.Validate(profile),
        };
    }

    var defaultProfilePath = Path.Combine(AppContext.BaseDirectory, "profiles", "default.pipebom.json");
    return store.LoadEffectiveProfile(
        options.AssemblyPath,
        new ProfileStoreOptions
        {
            DefaultProfilePath = defaultProfilePath,
            UserProfileDirectory = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "AFCA",
                "SolidWorksBOMAddin",
                "profiles"),
            CompanyProfileDirectory = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData),
                "AFCA",
                "SolidWorksBOMAddin",
                "profiles"),
        });
}

static IBomExporter CreateExporter(string format)
{
    return string.Equals(format, "xlsx", StringComparison.OrdinalIgnoreCase)
        ? new XlsxBomExporter()
        : new CsvBomExporter();
}

static void ExportDebugReportIfRequested(
    LauncherOptions options,
    ProfileLoadResult profileLoadResult,
    AssemblyReadResult assemblyReadResult,
    BomResult bomResult,
    IReadOnlyList<BomDiagnostic> diagnostics)
{
    if (string.IsNullOrWhiteSpace(options.DebugReportPath))
    {
        return;
    }

    var debugReport = new DebugReportService().Create(
        new DebugReportInput
        {
            AssemblyPath = assemblyReadResult.AssemblyPath ?? options.AssemblyPath,
            ProfilePath = profileLoadResult.SourcePath,
            Components = assemblyReadResult.Components,
            ComponentsScanned = assemblyReadResult.ComponentsScanned,
            ComponentsSkipped = assemblyReadResult.ComponentsSkipped,
            Rows = bomResult.Rows,
            Diagnostics = diagnostics,
        });

    var debugReportPath = options.DebugReportPath!;
    var directory = Path.GetDirectoryName(debugReportPath);
    if (!string.IsNullOrWhiteSpace(directory))
    {
        Directory.CreateDirectory(directory);
    }

    using var outputStream = File.Create(debugReportPath);
    new DebugReportExporter().Export(
        debugReport,
        outputStream,
        DebugReportExporter.GetFormatFromPath(debugReportPath));

    Console.WriteLine($"Exported debug report to {debugReportPath}");
}

static string ResolveOutputPath(LauncherOptions options)
{
    if (!string.IsNullOrWhiteSpace(options.OutputPath))
    {
        return options.OutputPath!;
    }

    var directory = Path.GetDirectoryName(options.AssemblyPath) ?? System.Environment.CurrentDirectory;
    var baseName = Path.GetFileNameWithoutExtension(options.AssemblyPath);
    var extension = string.Equals(options.Format, "xlsx", StringComparison.OrdinalIgnoreCase) ? ".xlsx" : ".csv";
    return Path.Combine(directory, $"{baseName}.bom{extension}");
}
