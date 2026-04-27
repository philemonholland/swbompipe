using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EPDM.Interop.epdm;

namespace BomPipePdmVaultInstaller;

internal static class Program
{
    private const string AddinName = "AFCA BOMPipe PDM";
    private const string AddinClassId = "{C68A6A86-9A42-42B4-96C7-A3CC31FD2C08}";
    private const string AddinDirectoryName = "BomPipePdmAddin";
    private const string AddinAssemblyName = "BomPipePdmAddin.dll";

    private static int Main(string[] args)
    {
        try
        {
            var options = VaultInstallerOptions.Parse(args);
            BomPipeVaultInstallerLog.Info($"Command '{options.Command}' started.");

            var exitCode = options.Command switch
            {
                "register" => RegisterAddin(options),
                "unregister" => UnregisterAddin(options),
                "list-vaults" => ListVaults(),
                "verify" => VerifyAddin(options),
                _ => throw new InvalidOperationException($"Unsupported command '{options.Command}'.")
            };
            BomPipeVaultInstallerLog.Info($"Command '{options.Command}' completed with exit code {exitCode}.");
            return exitCode;
        }
        catch (Exception ex)
        {
            BomPipeVaultInstallerLog.Error("Vault installer failed.", ex);
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }

    private static int RegisterAddin(VaultInstallerOptions options)
    {
        var payloadFiles = GetPayloadFiles(options.InstallRoot);
        var vaultNames = GetTargetVaultNames(options).ToArray();

        if (vaultNames.Length == 0)
        {
            Console.WriteLine("No local PDM vault views were found. Skipping PDM add-in registration.");
            return 0;
        }

        foreach (var vaultName in vaultNames)
        {
            var vault = new EdmVault5();
            vault.LoginAuto(vaultName, 0);
            BomPipeVaultInstallerLog.Info($"Registering PDM add-in in vault '{vaultName}'.");

            var addinManager = (IEdmAddInMgr5)vault;
            object result = string.Empty;
            addinManager.AddAddIns(
                string.Join("\n", payloadFiles),
                (int)(EdmAddAddInFlags.EdmAddin_AddAllFilesToOneAddIn | EdmAddAddInFlags.EdmAddin_ReplaceDuplicates),
                ref result);

            var installedAddin = GetInstalledAddin((IEdmAddInMgr9)vault);
            if (installedAddin is null)
            {
                throw new InvalidOperationException($"The BOMPipe PDM add-in was not discoverable in vault '{vaultName}' after registration.");
            }

            Console.WriteLine($"Registered {AddinName} in vault '{vaultName}'.");
            BomPipeVaultInstallerLog.Info($"Registered {AddinName} in vault '{vaultName}'.");
        }

        return 0;
    }

    private static int VerifyAddin(VaultInstallerOptions options)
    {
        var vaultNames = GetTargetVaultNames(options).ToArray();

        if (vaultNames.Length == 0)
        {
            Console.WriteLine("No local PDM vault views were found. Skipping PDM add-in verification.");
            return 0;
        }

        foreach (var vaultName in vaultNames)
        {
            var vault = new EdmVault5();
            vault.LoginAuto(vaultName, 0);
            BomPipeVaultInstallerLog.Info($"Verifying PDM add-in in vault '{vaultName}'.");

            var installedAddin = GetInstalledAddin((IEdmAddInMgr9)vault);
            if (installedAddin is null)
            {
                throw new InvalidOperationException($"The BOMPipe PDM add-in is not installed in vault '{vaultName}'.");
            }

            Console.WriteLine($"Verified {AddinName} in vault '{vaultName}'.");
            BomPipeVaultInstallerLog.Info($"Verified {AddinName} in vault '{vaultName}'.");
        }

        return 0;
    }

    private static int UnregisterAddin(VaultInstallerOptions options)
    {
        var vaultNames = GetTargetVaultNames(options).ToArray();

        if (vaultNames.Length == 0)
        {
            Console.WriteLine("No local PDM vault views were found. Skipping PDM add-in removal.");
            return 0;
        }

        foreach (var vaultName in vaultNames)
        {
            var vault = new EdmVault5();
            vault.LoginAuto(vaultName, 0);
            BomPipeVaultInstallerLog.Info($"Unregistering PDM add-in from vault '{vaultName}'.");

            var addinManager = (IEdmAddInMgr9)vault;
            var installedAddin = GetInstalledAddin(addinManager);

            if (installedAddin is null)
            {
                Console.WriteLine($"No BOMPipe PDM add-in was installed in vault '{vaultName}'.");
                continue;
            }

            RemoveInstalledAddin(addinManager, installedAddin.Value);
            Console.WriteLine($"Removed {AddinName} from vault '{vaultName}'.");
        }

        return 0;
    }

    private static int ListVaults()
    {
        foreach (var vaultName in GetRegisteredVaultNames())
        {
            Console.WriteLine(vaultName);
        }

        return 0;
    }

    private static IEnumerable<string> GetTargetVaultNames(VaultInstallerOptions options)
    {
        return options.VaultNames.Count > 0
            ? options.VaultNames
            : GetRegisteredVaultNames();
    }

    private static IEnumerable<string> GetRegisteredVaultNames()
    {
        var vault = new EdmVault5();
        var vault8 = (IEdmVault8)vault;
        EdmViewInfo[] views = Array.Empty<EdmViewInfo>();
        vault8.GetVaultViews(out views, false);

        return views
            .Select(view => view.mbsVaultName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string[] GetPayloadFiles(string installRoot)
    {
        var payloadDirectory = Path.Combine(installRoot, AddinDirectoryName);
        if (!Directory.Exists(payloadDirectory))
        {
            throw new DirectoryNotFoundException($"PDM add-in payload directory was not found at '{payloadDirectory}'.");
        }

        var primaryAssemblyPath = Path.Combine(payloadDirectory, AddinAssemblyName);
        if (!File.Exists(primaryAssemblyPath))
        {
            throw new FileNotFoundException($"PDM add-in assembly was not found at '{primaryAssemblyPath}'.", primaryAssemblyPath);
        }

        var dependencyFiles = Directory
            .GetFiles(payloadDirectory, "*.dll", SearchOption.TopDirectoryOnly)
            .Where(path => !path.Equals(primaryAssemblyPath, StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);

        return new[] { primaryAssemblyPath }.Concat(dependencyFiles).ToArray();
    }

    private static string NormalizeGuid(string? value)
    {
        return (value ?? string.Empty).Trim().Trim('{', '}');
    }

    private static EdmAddInInfo2? GetInstalledAddin(IEdmAddInMgr9 addinManager)
    {
        EdmAddInInfo2[] installedAddins = Array.Empty<EdmAddInInfo2>();
        addinManager.GetInstalledAddIns(out installedAddins);

        foreach (var installedAddin in installedAddins)
        {
            if (string.Equals(NormalizeGuid(installedAddin.mbsClassID), NormalizeGuid(AddinClassId), StringComparison.OrdinalIgnoreCase) ||
                string.Equals(installedAddin.mbsAddInName, AddinName, StringComparison.OrdinalIgnoreCase))
            {
                return installedAddin;
            }
        }

        return null;
    }

    private static void RemoveInstalledAddin(IEdmAddInMgr9 addinManager, EdmAddInInfo2 installedAddin)
    {
        var removeCandidates = new[]
        {
            installedAddin.mbsAddInName,
            installedAddin.mbsClassID,
            NormalizeGuid(installedAddin.mbsClassID),
            installedAddin.mbsModulePath
        }
        .Where(candidate => !string.IsNullOrWhiteSpace(candidate))
        .Distinct(StringComparer.OrdinalIgnoreCase);

        Exception? lastException = null;

        foreach (var candidate in removeCandidates)
        {
            try
            {
                addinManager.RemoveAddIn(candidate);
                return;
            }
            catch (Exception ex)
            {
                lastException = ex;
            }
        }

        throw new InvalidOperationException(
            $"Failed to remove the installed PDM add-in '{installedAddin.mbsAddInName}'.",
            lastException);
    }

    private sealed class VaultInstallerOptions
    {
        public string Command { get; set; } = string.Empty;
        public string InstallRoot { get; set; } = string.Empty;
        public List<string> VaultNames { get; set; } = new List<string>();

        public static VaultInstallerOptions Parse(string[] args)
        {
            if (args.Length == 0)
            {
                throw new InvalidOperationException(
                    "Usage: BomPipePdmVaultInstaller <register|unregister|verify|list-vaults> --install-root <path> [--vault <vault name>]");
            }

            var command = args[0].Trim().ToLowerInvariant();
            string? installRoot = null;
            var vaultNames = new List<string>();

            for (var index = 1; index < args.Length; index++)
            {
                var argument = args[index];
                switch (argument)
                {
                    case "--install-root":
                        installRoot = GetRequiredValue(args, ref index, argument);
                        break;
                    case "--vault":
                        vaultNames.Add(GetRequiredValue(args, ref index, argument));
                        break;
                    default:
                        throw new InvalidOperationException($"Unknown argument '{argument}'.");
                }
            }

            if (!string.Equals(command, "list-vaults", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(installRoot))
            {
                throw new InvalidOperationException("--install-root is required for register, unregister, and verify.");
            }

            return new VaultInstallerOptions
            {
                Command = command,
                InstallRoot = installRoot ?? string.Empty,
                VaultNames = vaultNames
            };
        }

        private static string GetRequiredValue(string[] args, ref int index, string argumentName)
        {
            if (index + 1 >= args.Length)
            {
                throw new InvalidOperationException($"Missing value for '{argumentName}'.");
            }

            index++;
            return args[index];
        }
    }
}
