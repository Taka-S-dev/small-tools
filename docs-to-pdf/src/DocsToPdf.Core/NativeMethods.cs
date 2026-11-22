using System.Runtime.InteropServices;

namespace DocsToPdf.Core;

internal static class NativeMethods
{
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int StrCmpLogicalW(string x, string y);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    public static extern int GetWindowLong(IntPtr hWnd, int index);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint GetConsoleProcessList(uint[] processList, uint processCount);

    [DllImport("ole32.dll")]
    public static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}

/// <summary>エクスプローラーと同じ並び（"2" が "10" より前）で文字列を比べる。</summary>
internal sealed class NaturalStringComparer : IComparer<string>
{
    public static readonly NaturalStringComparer Instance = new();

    public int Compare(string? x, string? y)
    {
        if (x is null) return y is null ? 0 : -1;
        if (y is null) return 1;
        return NativeMethods.StrCmpLogicalW(x, y);
    }
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

/// <summary>
/// Office が処理中で呼び出しを断ったとき（RPC_E_SERVERCALL_RETRYLATER）に、
/// 例外にせず少し待って再試行させる。登録したスレッド（STA）でだけ効く。
/// </summary>
[ComVisible(true)]
internal sealed class OleMessageFilter : IOleMessageFilter, IDisposable
{
    private const int ServerCallRetryLater = 2;
    private const int RetryDelayMs = 250;
    private const int GiveUpAfterMs = 120_000;

    private readonly IOleMessageFilter? _previous;

    private OleMessageFilter()
    {
        NativeMethods.CoRegisterMessageFilter(this, out _previous);
    }

    public static OleMessageFilter Register() => new();

    public void Dispose() => NativeMethods.CoRegisterMessageFilter(_previous, out _);

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        => 0; // SERVERCALL_ISHANDLED

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        => dwRejectType == ServerCallRetryLater && dwTickCount < GiveUpAfterMs ? RetryDelayMs : -1;

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        => 2; // PENDINGMSG_WAITDEFPROCESS
}
