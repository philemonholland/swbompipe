using System.Windows.Forms;

namespace SolidWorksBOMAddin;

internal sealed class SolidWorksKeyboardMessageFilter : IMessageFilter
{
    private const int WmKeyDown = 0x0100;
    private const int WmKeyUp = 0x0101;
    private const int WmChar = 0x0102;
    private const int WmDeadChar = 0x0103;
    private const int WmSysKeyDown = 0x0104;
    private const int WmSysKeyUp = 0x0105;
    private const int WmSysChar = 0x0106;
    private const int WmSysDeadChar = 0x0107;
    private const int WmUniChar = 0x0109;

    private readonly BomPreviewShellForm _form;

    public SolidWorksKeyboardMessageFilter(BomPreviewShellForm form)
    {
        _form = form ?? throw new ArgumentNullException(nameof(form));
    }

    public bool PreFilterMessage(ref Message m)
    {
        return IsKeyboardMessage(m.Msg) && _form.ShouldTrapExternalKeyboardMessage(m.HWnd);
    }

    internal static bool IsKeyboardMessage(int message)
    {
        return message is WmKeyDown
            or WmKeyUp
            or WmChar
            or WmDeadChar
            or WmSysKeyDown
            or WmSysKeyUp
            or WmSysChar
            or WmSysDeadChar
            or WmUniChar;
    }
}
