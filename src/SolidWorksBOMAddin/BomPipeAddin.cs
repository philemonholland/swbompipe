using System.Runtime.InteropServices;
using BomCore;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SolidWorksBOMAddin;

[ComVisible(true)]
[Guid("E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A")]
[ProgId("AFCA.PipingBom.Generator")]
public sealed class BomPipeAddin : ISwAddin
{
    private const int PreviewCommandGroupId = 0x5A01;

    private ISldWorks? _application;
    private int _cookie;
    private CommandManager? _commandManager;
    private CommandGroup? _commandGroup;
    private BomPreviewShellForm? _previewShell;

    public string AddinName => "AFCA Piping BOM Generator";

    public IAssemblyReader? AssemblyReader => _application is null ? null : new SolidWorksAssemblyReader(_application);

    public ISelectedComponentPropertyReader? SelectedComponentPropertyReader =>
        _application is null ? null : new SolidWorksSelectedComponentPropertyReader(_application);

    public bool ConnectToSW(object thisSw, int cookie)
    {
        _application = thisSw as ISldWorks
            ?? throw new InvalidCastException("SolidWorks did not provide an ISldWorks application object.");
        _cookie = cookie;
        RegisterCommandGroup();
        return true;
    }

    public bool DisconnectFromSW()
    {
        if (_previewShell is not null)
        {
            var shell = _previewShell;
            _previewShell = null;
            shell.Close();
            shell.Dispose();
        }

        if (_commandManager is not null)
        {
            try
            {
                _commandManager.RemoveCommandGroup2(PreviewCommandGroupId, true);
            }
            catch
            {
                // Best effort cleanup when SolidWorks is shutting down.
            }
        }

        _commandGroup = null;
        _commandManager = null;
        _application = null;
        _cookie = 0;
        return true;
    }

    public void OpenPreviewShell()
    {
        if (_previewShell is null || _previewShell.IsDisposed)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            _previewShell = new BomPreviewShellForm(this);
            _previewShell.FormClosed += (_, _) =>
            {
                if (_previewShell?.IsDisposed == true)
                {
                    _previewShell = null;
                }
            };
            _previewShell.Show();
        }
        else
        {
            if (_previewShell.WindowState == System.Windows.Forms.FormWindowState.Minimized)
            {
                _previewShell.WindowState = System.Windows.Forms.FormWindowState.Normal;
            }

            _previewShell.BringToFront();
            _previewShell.Focus();
        }
    }

    public int CanOpenPreviewShell()
    {
        return 1;
    }

    internal ISldWorks RequireApplication()
    {
        return _application ?? throw new InvalidOperationException("The add-in is not connected to SolidWorks.");
    }

    private void RegisterCommandGroup()
    {
        var application = RequireApplication();
        application.SetAddinCallbackInfo2(0L, this, _cookie);
        _commandManager = application.GetCommandManager(_cookie);

        var errors = 0;
        _commandGroup = _commandManager.CreateCommandGroup2(
            PreviewCommandGroupId,
            "Pipe BOM",
            "Open the pipe BOM preview and mapping shell.",
            string.Empty,
            -1,
            true,
            ref errors);

        if (_commandGroup is null)
        {
            throw new InvalidOperationException($"SolidWorks did not create the Pipe BOM command group (error {errors}).");
        }

        _commandGroup.HasMenu = true;
        _commandGroup.HasToolbar = false;
        _commandGroup.AddCommandItem2(
            "Open BOM Preview Shell",
            -1,
            "Open the minimal preview and mapping shell.",
            "Open BOM Preview Shell",
            -1,
            nameof(OpenPreviewShell),
            nameof(CanOpenPreviewShell),
            0,
            (int)swCommandItemType_e.swMenuItem);
        _commandGroup.Activate();
    }
}
