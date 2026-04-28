using System.Runtime.InteropServices;

namespace SolidWorksBOMAddin;

internal sealed class SolidWorksPreTranslateKeyboardHook : IDisposable
{
    private const int WhGetMessage = 3;
    private const int PmRemove = 0x0001;
    private const int WmNull = 0x0000;
    private const int WmKeyDown = 0x0100;
    private const int WmSysKeyDown = 0x0104;

    private readonly BomPreviewShellForm _form;
    private readonly HookProc _hookProc;
    private IntPtr _hookHandle;

    public SolidWorksPreTranslateKeyboardHook(BomPreviewShellForm form)
    {
        _form = form ?? throw new ArgumentNullException(nameof(form));
        _hookProc = HookCallback;
    }

    public void Install()
    {
        if (_hookHandle != IntPtr.Zero)
        {
            return;
        }

        _hookHandle = SetWindowsHookEx(WhGetMessage, _hookProc, IntPtr.Zero, GetCurrentThreadId());
        if (_hookHandle == IntPtr.Zero)
        {
            throw new InvalidOperationException($"Could not install the BPM keyboard hook (Win32 error {Marshal.GetLastWin32Error()}).");
        }
    }

    public void Dispose()
    {
        if (_hookHandle == IntPtr.Zero)
        {
            return;
        }

        UnhookWindowsHookEx(_hookHandle);
        _hookHandle = IntPtr.Zero;
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code >= 0 && wParam == (IntPtr)PmRemove)
        {
            var message = Marshal.PtrToStructure<NativeMessage>(lParam);
            if ((message.Message == WmKeyDown || message.Message == WmSysKeyDown)
                && _form.TryInterceptSingleLetterShortcut(message.WParam, message.LParam))
            {
                message.Message = WmNull;
                message.WParam = IntPtr.Zero;
                message.LParam = IntPtr.Zero;
                Marshal.StructureToPtr(message, lParam, false);
            }
        }

        return CallNextHookEx(_hookHandle, code, wParam, lParam);
    }

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public IntPtr HWnd;
        public int Message;
        public IntPtr WParam;
        public IntPtr LParam;
        public int Time;
        public int PtX;
        public int PtY;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
