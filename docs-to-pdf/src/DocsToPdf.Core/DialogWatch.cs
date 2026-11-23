using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace DocsToPdf.Core;

/// <summary>
/// このツールが起動した Office が、確認の画面を出したまま答えを待って止まっていないかを見張る。
///
/// 確認を出さない設定にしていても、Office が画面を出して待つことがある（想定していない種類の確認など）。
/// 時間切れ（既定 5 分）まで待つと、その間ほかのファイルが進まないうえ、利用者の画面にその確認が出たままになる。
/// そこで「確認の画面が出ている」かつ「その間 Office が何もしていない（CPU をほとんど使っていない）」状態が
/// しばらく続いたら止まったと判断し、画面の文面を添えて知らせる。
/// 大きな文書を書き出している間の進み具合の画面は、Office が働いているので止まったとは見なさない。
/// </summary>
internal sealed class DialogWatch : IDisposable
{
    /// <summary>確認の画面が出たまま、これだけ何もしなければ止まったと見なす。</summary>
    public static readonly TimeSpan Grace = TimeSpan.FromSeconds(5);
    /// <summary>この間に使った CPU 時間がこれより少なければ「何もしていない」。</summary>
    public static readonly TimeSpan IdleCpu = TimeSpan.FromMilliseconds(250);
    private static readonly TimeSpan Interval = TimeSpan.FromSeconds(1);

    private readonly Process _process;
    private readonly Action<string> _onStuck;
    private readonly Timer _timer;
    private readonly Judge _judge = new();
    private int _ticking;

    public DialogWatch(Process process, Action<string> onStuck)
    {
        _process = process;
        _onStuck = onStuck;
        _timer = new Timer(_ => Tick(), null, Interval, Interval);
    }

    private void Tick()
    {
        if (Interlocked.Exchange(ref _ticking, 1) == 1) return; // 前の確認がまだ終わっていない
        try
        {
            _process.Refresh();
            if (_process.HasExited) return;
            IntPtr dialog = FindDialog(_process.Id);
            if (_judge.Step(dialog, _process.TotalProcessorTime, DateTime.UtcNow))
                _onStuck(Describe(dialog));
        }
        catch { /* プロセスが終わった途中など。次の回に見直す */ }
        finally
        {
            Volatile.Write(ref _ticking, 0);
        }
    }

    public void Dispose() => _timer.Dispose();

    /// <summary>
    /// 判定の部分（画面の操作から切り離して確かめられるようにしてある）。
    /// 同じ確認の画面が出続け、その間に使った CPU 時間が IdleCpu 未満のまま Grace を過ぎたら、1 度だけ true を返す。
    /// </summary>
    internal sealed class Judge
    {
        private IntPtr _dialog;
        private DateTime _since;
        private TimeSpan _cpuSince;
        private bool _fired;

        public bool Step(IntPtr dialog, TimeSpan cpu, DateTime now)
        {
            if (_fired) return false;
            if (dialog == IntPtr.Zero)
            {
                _dialog = IntPtr.Zero;
                return false;
            }
            // 新しい画面が出た、または Office が働いている間は、数え直す
            if (dialog != _dialog || cpu - _cpuSince >= IdleCpu)
            {
                (_dialog, _since, _cpuSince) = (dialog, now, cpu);
                return false;
            }
            if (now - _since < Grace) return false;
            _fired = true;
            return true;
        }
    }

    /// <summary>そのプロセスが出している、見えている確認の画面（なければ IntPtr.Zero）。</summary>
    internal static IntPtr FindDialog(int processId) => FindVisibleWindow(processId, IsDialogClass);

    /// <summary>そのプロセスが出している、見えている窓のうち、クラス名が合う最初のもの（なければ IntPtr.Zero）。</summary>
    internal static IntPtr FindVisibleWindow(int processId, Func<string, bool> classMatches)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid != processId || !IsWindowVisible(hwnd)) return true;
            if (!classMatches(ClassName(hwnd))) return true;
            found = hwnd;
            return false;
        }, IntPtr.Zero);
        return found;
    }

    /// <summary>
    /// 確認の画面のクラス名か。#32770 は Windows 標準のダイアログとメッセージ ボックス、
    /// bosa_sdm_ で始まるものは Office 独自のダイアログ、NUIDialog は新しい Office のダイアログ。
    /// </summary>
    internal static bool IsDialogClass(string className) =>
        className == "#32770"
        || className.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase)
        || className.Equals("NUIDialog", StringComparison.OrdinalIgnoreCase);

    /// <summary>画面の題名と、中の文（ボタン以外）をつないだもの。Office 独自のダイアログは中の文を読めないことがある。</summary>
    internal static string Describe(IntPtr dialog)
    {
        var parts = new List<string>();
        string title = WindowText(dialog);
        if (title.Length > 0) parts.Add(title);
        EnumChildWindows(dialog, (child, _) =>
        {
            if (!ClassName(child).Equals("Button", StringComparison.OrdinalIgnoreCase))
            {
                string text = WindowText(child).Replace("\r", " ").Replace("\n", " ").Trim();
                if (text.Length > 0 && !parts.Contains(text)) parts.Add(text);
            }
            return true;
        }, IntPtr.Zero);
        string joined = string.Join(" / ", parts);
        return joined.Length > 200 ? joined[..200] + "…" : joined;
    }

    private static string ClassName(IntPtr hwnd)
    {
        var sb = new StringBuilder(256);
        return GetClassName(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : "";
    }

    private static string WindowText(IntPtr hwnd)
    {
        int length = GetWindowTextLength(hwnd);
        if (length <= 0) return "";
        var sb = new StringBuilder(length + 1);
        return GetWindowText(hwnd, sb, sb.Capacity) > 0 ? sb.ToString() : "";
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hwnd);
}
