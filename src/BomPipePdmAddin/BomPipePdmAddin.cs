using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using EPDM.Interop.epdm;

namespace BomPipePdmAddin;

[ComVisible(true)]
[Guid("C68A6A86-9A42-42B4-96C7-A3CC31FD2C08")]
[ProgId("AFCA.BOMPipe.PdmAddin")]
[ClassInterface(ClassInterfaceType.None)]
public sealed class BomPipePdmAddin : IEdmAddIn5
{
    private const int GenerateBomCommandId = 1;
    private const string AddinName = "AFCA BOMPipe PDM";
    private const string CompanyName = "AFCA";
    private const string Description = "Generate BOMPipe exports for selected SolidWorks assemblies.";
    private const string MenuText = "Generate BOM with BOMPipe";
    private const string InvokerScriptName = "Invoke-BOMPipe.ps1";

    public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
    {
        BomPipePdmLog.Info("GetAddInInfo entered.");
        poInfo.mbsAddInName = AddinName;
        poInfo.mbsCompany = CompanyName;
        poInfo.mbsDescription = Description;
        poInfo.mlAddInVersion = 1;
        poInfo.mlRequiredVersionMajor = 17;
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_Menu);

        var menuFlags =
            (int)EdmMenuFlags.EdmMenu_MustHaveSelection |
            (int)EdmMenuFlags.EdmMenu_OnlyFiles |
            (int)EdmMenuFlags.EdmMenu_OnlySingleSelection |
            (int)EdmMenuFlags.EdmMenu_OnlyInContextMenu |
            (int)EdmMenuFlags.EdmMenu_ContextMenuItem;

        poCmdMgr.AddCmd(
            GenerateBomCommandId,
            MenuText,
            menuFlags,
            "Generate a BOMPipe export for the selected SolidWorks assembly.",
            "Generate BOM with BOMPipe",
            0,
            0);
        BomPipePdmLog.Info("PDM menu command registered.");
    }

    public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
    {
        BomPipePdmLog.Info($"OnCmd entered. Type={poCmd.meCmdType}; Id={poCmd.mlCmdID}.");
        if (poCmd.meCmdType != EdmCmdType.EdmCmd_Menu || poCmd.mlCmdID != GenerateBomCommandId)
        {
            return;
        }

        try
        {
            var vault = poCmd.mpoVault as IEdmVault5;
            var selection = ppoData;

            if (vault is null || selection is null || selection.Length == 0)
            {
                BomPipePdmLog.Info("No PDM file selection received.");
                ShowMessage(vault, "BOMPipe did not receive a selected file from PDM.");
                return;
            }

            var assemblyPath = ResolveAssemblyPath(vault, selection[0]);
            BomPipePdmLog.Info($"Resolved PDM selection path: {assemblyPath ?? "<null>"}");
            if (string.IsNullOrWhiteSpace(assemblyPath))
            {
                ShowMessage(vault, "BOMPipe could not resolve the selected vault file to a local path.");
                return;
            }

            if (!assemblyPath!.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage(vault, "BOMPipe only supports SolidWorks assembly (*.SLDASM) selections.");
                return;
            }

            if (!File.Exists(assemblyPath))
            {
                ShowMessage(vault, $"The selected assembly is not available locally:{Environment.NewLine}{assemblyPath}");
                return;
            }

            var invokerPath = GetInvokerPath();
            if (!File.Exists(invokerPath))
            {
                BomPipePdmLog.Info($"Invoker script missing at {invokerPath}.");
                ShowMessage(
                    vault,
                    "BOMPipe is not installed in the current user profile. Run install-bompipe.cmd first.");
                return;
            }

            var processStartInfo = new ProcessStartInfo
            {
                FileName = GetWindowsPowerShellPath(),
                Arguments =
                    $"-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{invokerPath}\" \"{assemblyPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(invokerPath) ?? GetInstallRoot()
            };

            Process.Start(processStartInfo);
            BomPipePdmLog.Info("BOMPipe invoker process started.");
        }
        catch (Exception ex)
        {
            BomPipePdmLog.Error("OnCmd failed.", ex);
            ShowMessage(poCmd.mpoVault as IEdmVault5, $"BOMPipe failed to launch:{Environment.NewLine}{ex.Message}");
        }
    }

    private static string? ResolveAssemblyPath(IEdmVault5 vault, EdmCmdData commandData)
    {
        if (commandData.mlObjectID1 == 0 || commandData.mlObjectID2 == 0)
        {
            return commandData.mbsStrData1;
        }

        var file = vault.GetObject(EdmObjectType.EdmObject_File, commandData.mlObjectID1) as IEdmFile5;
        var folder = vault.GetObject(EdmObjectType.EdmObject_Folder, commandData.mlObjectID2) as IEdmFolder5;

        if (file is null || folder is null)
        {
            return commandData.mbsStrData1;
        }

        return Path.Combine(folder.LocalPath, file.Name);
    }

    private static string GetInstallRoot()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AFCA",
            "BOMPipe");
    }

    private static string GetInvokerPath()
    {
        return Path.Combine(GetInstallRoot(), InvokerScriptName);
    }

    private static string GetWindowsPowerShellPath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            "System32",
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");
    }

    private static void ShowMessage(IEdmVault5? vault, string message)
    {
        if (vault is IEdmVault10 vault10)
        {
            vault10.MsgBox(0, message);
        }
    }
}
