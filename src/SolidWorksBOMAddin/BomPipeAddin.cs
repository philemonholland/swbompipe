using System.Runtime.InteropServices;
using System.Diagnostics;
using BomCore;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace SolidWorksBOMAddin;

[ComVisible(true)]
[Guid("E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A")]
[ProgId("AFCA.PipingBom.Generator")]
[ClassInterface(ClassInterfaceType.AutoDual)]
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
        try
        {
            BomPipeLog.Info("ConnectToSW entered.");
            _application = thisSw as ISldWorks
                ?? throw new InvalidCastException("SolidWorks did not provide an ISldWorks application object.");
            _cookie = cookie;
            RegisterCommandGroup();
            BomPipeLog.Info("ConnectToSW completed.");
            return true;
        }
        catch (Exception ex)
        {
            BomPipeLog.Error("ConnectToSW failed.", ex);
            ShowSolidWorksError("BOMPipe failed to load.", ex);
            return false;
        }
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
        try
        {
            BomPipeLog.Info("OpenPreviewShell entered.");
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

                // Keep BPM modeless, but do not show it as an owned SolidWorks window.
                // SolidWorks can continue to process some accelerator keys through the owner relationship.
                _previewShell.Show();
                _previewShell.EnsureKeyboardActivation();
            }
            else
            {
                if (_previewShell.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                {
                    _previewShell.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }

                _previewShell.Show();
                _previewShell.EnsureKeyboardActivation();
            }

            BomPipeLog.Info("OpenPreviewShell completed.");
        }
        catch (Exception ex)
        {
            BomPipeLog.Error("OpenPreviewShell failed.", ex);
            ShowSolidWorksError("BOMPipe could not open the preview shell.", ex);
        }
    }

    public int CanOpenPreviewShell()
    {
        BomPipeLog.Info("CanOpenPreviewShell queried.");
        return 1;
    }

    internal ISldWorks RequireApplication()
    {
        return _application ?? throw new InvalidOperationException("The add-in is not connected to SolidWorks.");
    }

    internal IntPtr GetSolidWorksMainWindowHandle()
    {
        return Process.GetCurrentProcess().MainWindowHandle;
    }

    private void RegisterCommandGroup()
    {
        var application = RequireApplication();
        BomPipeLog.Info("Registering SolidWorks command group.");
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
        BomPipeLog.Info("SolidWorks command group registered.");
    }

    private System.Windows.Forms.IWin32Window? GetSolidWorksOwnerWindow()
    {
        try
        {
            dynamic application = RequireApplication();
            dynamic frame = application.Frame();
            var handle = Convert.ToInt32(frame.GetHWnd());
            return handle == 0 ? null : new NativeWindowWrapper(new IntPtr(handle));
        }
        catch (Exception ex)
        {
            BomPipeLog.Error("Could not resolve SolidWorks owner window.", ex);
            return null;
        }
    }

    private void ShowSolidWorksError(string message, Exception exception)
    {
        try
        {
            _application?.SendMsgToUser2(
                $"{message}{System.Environment.NewLine}{exception.Message}",
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk);
        }
        catch
        {
            // SolidWorks may be shutting down or unavailable.
        }
    }

    private sealed class NativeWindowWrapper : System.Windows.Forms.IWin32Window
    {
        public NativeWindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }
    }
}
