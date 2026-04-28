using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SolidWorksBOMAddin;

internal static class NativeWindowActivation
{
    private const int Show = 5;
    private const int Restore = 9;
    private static readonly IntPtr HwndTopMost = new(-1);
    private static readonly IntPtr HwndNoTopMost = new(-2);
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoActivate = 0x0010;

    public static void Activate(Form form, Control? preferredControl = null)
    {
        ArgumentNullException.ThrowIfNull(form);

        if (form.IsDisposed || !form.IsHandleCreated)
        {
            return;
        }

        var formHandle = form.Handle;
        var targetHandle = preferredControl is not null && preferredControl.IsHandleCreated
            ? preferredControl.Handle
            : formHandle;
        var foregroundHandle = GetForegroundWindow();
        if (foregroundHandle == formHandle)
        {
            SetFocus(targetHandle);
            return;
        }

        var currentThreadId = GetCurrentThreadId();
        var foregroundThreadId = foregroundHandle == IntPtr.Zero
            ? 0U
            : GetWindowThreadProcessId(foregroundHandle, out _);
        var attached = false;

        try
        {
            if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
            {
                attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);
            }

            ShowWindow(formHandle, form.WindowState == FormWindowState.Minimized ? Restore : Show);
            SetWindowPos(formHandle, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize);
            SetWindowPos(formHandle, HwndNoTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate);
            BringWindowToTop(formHandle);
            SetActiveWindow(formHandle);
            SetForegroundWindow(formHandle);
            SetFocus(targetHandle);
        }
        finally
        {
            if (attached)
            {
                AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }
        }
    }

    internal static IntPtr GetFocusedWindow()
    {
        return GetFocus();
    }

    internal static bool IsWindowOrDescendant(IntPtr ancestorHandle, IntPtr candidateHandle)
    {
        return ancestorHandle != IntPtr.Zero
            && candidateHandle != IntPtr.Zero
            && (ancestorHandle == candidateHandle || IsChild(ancestorHandle, candidateHandle));
    }

    internal static void PostCharacter(IntPtr handle, char character, IntPtr lParam)
    {
        if (handle == IntPtr.Zero)
        {
            return;
        }

        PostMessage(handle, 0x0102, (IntPtr)character, lParam);
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int command);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetActiveWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int x,
        int y,
        int cx,
        int cy,
        uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
