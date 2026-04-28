using System.Windows.Forms;

namespace SolidWorksBOMAddin;

internal sealed class SolidWorksMainWindowKeyboardInterceptor : NativeWindow, IDisposable
{
    private readonly BomPreviewShellForm _form;
    private IntPtr _attachedHandle;

    public SolidWorksMainWindowKeyboardInterceptor(BomPreviewShellForm form)
    {
        _form = form ?? throw new ArgumentNullException(nameof(form));
    }

    public void Attach(IntPtr handle)
    {
        if (handle == IntPtr.Zero || handle == _attachedHandle)
        {
            return;
        }

        Detach();
        AssignHandle(handle);
        _attachedHandle = handle;
    }

    public void Dispose()
    {
        Detach();
    }

    protected override void WndProc(ref Message m)
    {
        if (SolidWorksKeyboardMessageFilter.IsKeyboardMessage(m.Msg)
            && _form.ShouldTrapExternalKeyboardMessage(_attachedHandle))
        {
            return;
        }

        base.WndProc(ref m);
    }

    private void Detach()
    {
        if (_attachedHandle == IntPtr.Zero)
        {
            return;
        }

        ReleaseHandle();
        _attachedHandle = IntPtr.Zero;
    }
}
