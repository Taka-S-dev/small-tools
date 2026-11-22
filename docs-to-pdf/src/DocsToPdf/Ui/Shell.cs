using System.IO.Pipes;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace DocsToPdf.Ui;

/// <summary>拡張子ごとの、エクスプローラーと同じ小さいアイコン。</summary>
internal static class ShellIcons
{
    private const uint ShgfiIcon = 0x100, ShgfiSmallIcon = 0x1, ShgfiUseFileAttributes = 0x10;
    private const uint FileAttributeNormal = 0x80;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShFileInfo
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(string path, uint attributes, ref ShFileInfo info, uint size, uint flags);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr handle);

    /// <summary>その拡張子のアイコンを取る。ファイルの実体は見ないので、存在しなくてもよい。</summary>
    public static Icon? ForExtension(string extension)
    {
        var info = new ShFileInfo();
        var result = SHGetFileInfo("x" + extension, FileAttributeNormal, ref info, (uint)Marshal.SizeOf<ShFileInfo>(),
            ShgfiIcon | ShgfiSmallIcon | ShgfiUseFileAttributes);
        if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero) return null;
        try
        {
            return (Icon)Icon.FromHandle(info.hIcon).Clone();
        }
        finally
        {
            DestroyIcon(info.hIcon);
        }
    }
}

/// <summary>タスクバーのボタンに進み具合を出す。</summary>
internal sealed class TaskbarProgress
{
    [ComImport, Guid("ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface ITaskbarList3
    {
        void HrInit();
        void AddTab(IntPtr hwnd);
        void DeleteTab(IntPtr hwnd);
        void ActivateTab(IntPtr hwnd);
        void SetActiveAlt(IntPtr hwnd);
        void MarkFullscreenWindow(IntPtr hwnd, [MarshalAs(UnmanagedType.Bool)] bool fullscreen);
        void SetProgressValue(IntPtr hwnd, ulong completed, ulong total);
        void SetProgressState(IntPtr hwnd, int flags);
    }

    [ComImport, Guid("56FDF344-FD6D-11d0-958A-006097C9A090"), ClassInterface(ClassInterfaceType.None)]
    private class TaskbarInstance;

    private const int NoProgress = 0, Indeterminate = 1, Normal = 2;

    private readonly ITaskbarList3? _taskbar;

    public TaskbarProgress()
    {
        try
        {
            _taskbar = (ITaskbarList3)new TaskbarInstance();
            _taskbar.HrInit();
        }
        catch
        {
            _taskbar = null; // タスクバーがない環境でも動作は続ける
        }
    }

    public void Set(IntPtr window, int done, int total) => Try(() =>
    {
        _taskbar!.SetProgressState(window, Normal);
        _taskbar.SetProgressValue(window, (ulong)Math.Max(0, done), (ulong)Math.Max(1, total));
    });

    public void SetBusy(IntPtr window) => Try(() => _taskbar!.SetProgressState(window, Indeterminate));

    public void Clear(IntPtr window) => Try(() => _taskbar!.SetProgressState(window, NoProgress));

    private void Try(Action action)
    {
        if (_taskbar is null) return;
        try { action(); } catch { /* 表示だけの機能なので失敗は無視する */ }
    }
}

internal static class WindowAttention
{
    [StructLayout(LayoutKind.Sequential)]
    private struct FlashInfo
    {
        public uint cbSize;
        public IntPtr hwnd;
        public uint dwFlags;
        public uint uCount;
        public uint dwTimeout;
    }

    [DllImport("user32.dll")]
    private static extern bool FlashWindowEx(ref FlashInfo info);

    [DllImport("user32.dll")]
    private static extern bool AllowSetForegroundWindow(int processId);

    /// <summary>ウィンドウが前面にないとき、タスクバーのボタンを光らせて終わったことを知らせる。</summary>
    public static void Flash(Form form)
    {
        if (Form.ActiveForm == form) return;
        const uint All = 3, UntilForeground = 12;
        var info = new FlashInfo
        {
            cbSize = (uint)Marshal.SizeOf<FlashInfo>(),
            hwnd = form.Handle,
            dwFlags = All | UntilForeground,
            uCount = 3,
        };
        FlashWindowEx(ref info);
    }

    /// <summary>ほかのプロセス（先に起動している自分）が前面に出るのを許す。</summary>
    public static void AllowOthersToTakeForeground() => AllowSetForegroundWindow(-1);
}

/// <summary>
/// 二重起動を防ぐ。2 つ目を起動すると、渡されたファイルを先に起動しているウィンドウへ送って終わる。
/// 「送る」メニューや exe へのドロップを繰り返しても、1 つのリストに集まる。
/// </summary>
internal sealed class SingleInstance : IDisposable
{
    private readonly string _pipeName;
    private readonly Mutex _mutex;
    private readonly CancellationTokenSource _stop = new();

    private SingleInstance(Mutex mutex, string pipeName)
    {
        _mutex = mutex;
        _pipeName = pipeName;
    }

    /// <summary>最初の起動なら自分を返す。すでに起動していれば lines（絶対パスと画面の指定）を送って null を返す。</summary>
    public static SingleInstance? Acquire(string[] lines)
    {
        // パイプの名前はマシン全体で 1 つなので、利用者（SID）とセッションで分け、二重起動の判定（Local\ はセッションごと）とそろえる
        string user = WindowsIdentity.GetCurrent().User?.Value ?? Convert.ToHexString(Encoding.UTF8.GetBytes(Environment.UserName));
        using var self = Process.GetCurrentProcess();
        string name = $"DocsToPdf-{user}-{self.SessionId}";
        var mutex = new Mutex(initiallyOwned: true, $@"Local\{name}", out bool isFirst);
        if (isFirst) return new SingleInstance(mutex, name);

        mutex.Dispose();
        try
        {
            WindowAttention.AllowOthersToTakeForeground();
            // 同じ名前のパイプをほかの利用者が先に作っていても、ファイルのパスをそちらへ送らない
            using var client = new NamedPipeClientStream(".", name, PipeDirection.Out, PipeOptions.CurrentUserOnly);
            client.Connect(10_000);
            using var writer = new StreamWriter(client, new UTF8Encoding(false));
            foreach (var line in lines) writer.WriteLine(line);
        }
        catch
        {
            // 先の起動が応答しないときは、何もせず終わる
        }
        return null;
    }

    /// <summary>あとから起動されたときに送られてくる行（ファイルのパスと画面の指定）を受け取り始める。</summary>
    public void Listen(Action<string[]> onLines)
    {
        _ = Task.Run(async () =>
        {
            while (!_stop.IsCancellationRequested)
            {
                try
                {
                    using var server = new NamedPipeServerStream(_pipeName, PipeDirection.In, 1,
                        PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
                    await server.WaitForConnectionAsync(_stop.Token);
                    using var reader = new StreamReader(server, Encoding.UTF8);
                    string text = await reader.ReadToEndAsync(_stop.Token);
                    onLines(text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch
                {
                    await Task.Delay(500);
                }
            }
        });
    }

    public void Dispose()
    {
        _stop.Cancel();
        try { _mutex.ReleaseMutex(); } catch { /* 取得したスレッド以外からは解放できない */ }
        _mutex.Dispose();
    }
}
