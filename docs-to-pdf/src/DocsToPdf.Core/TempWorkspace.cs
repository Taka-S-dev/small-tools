using System.Diagnostics;

namespace DocsToPdf.Core;

/// <summary>
/// アプリが動いている間に使う作業フォルダ（%TEMP%\DocsToPdf\...）。変換した PDF を置く。Dispose で中身ごと消す。
/// 異常終了で残ったフォルダは、次に起動したときに片付ける。
/// </summary>
internal sealed class TempWorkspace : IDisposable
{
    private const string OwnerFile = "owner";
    private static readonly string Root = Path.Combine(Path.GetTempPath(), "DocsToPdf");

    private TempWorkspace(string folderPath) => FolderPath = folderPath;

    public string FolderPath { get; }

    public static TempWorkspace Create()
    {
        string path = Path.Combine(Root, $"{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        // このフォルダを使っているプロセスを書いておく（片付けのときに、まだ動いているかを見る）
        using var self = Process.GetCurrentProcess();
        File.WriteAllText(Path.Combine(path, OwnerFile), ProcessStamp.Format(self));
        return new TempWorkspace(path);
    }

    public void Dispose()
    {
        try { Directory.Delete(FolderPath, recursive: true); } catch { /* 次回起動時に消す */ }
    }

    /// <summary>
    /// 持ち主のプロセスがもう動いていない作業フォルダを片付ける。
    /// そのときに起動していた Office が見えないまま残っていれば、それも終了させる。
    /// </summary>
    public static void CleanupStale()
    {
        try
        {
            if (!Directory.Exists(Root)) return;
            foreach (var dir in Directory.EnumerateDirectories(Root))
            {
                try
                {
                    string owner = Path.Combine(dir, OwnerFile);
                    if (File.Exists(owner))
                    {
                        if (ProcessStamp.IsAlive(File.ReadAllText(owner))) continue;
                        OfficeLedger.KillLeftovers(dir);
                    }
                    else if (Directory.GetLastWriteTimeUtc(dir) > DateTime.UtcNow.AddDays(-1))
                    {
                        continue; // 持ち主が分からないものは、十分に古くなってから消す
                    }
                    Directory.Delete(dir, recursive: true);
                }
                catch { /* 使用中なら次回 */ }
            }
        }
        catch { /* 片付けに失敗しても起動は続ける */ }
    }
}

/// <summary>プロセスを「番号と起動時刻」で書き表す。番号は使い回されるので、起動時刻も合わせて同じプロセスかを確かめる。</summary>
internal static class ProcessStamp
{
    public static string Format(Process process) => $"{process.Id},{process.StartTime.ToUniversalTime().Ticks}";

    /// <summary>書かれたプロセスがまだ動いていれば、そのプロセスを返す。呼び出し側で Dispose すること。</summary>
    public static Process? Find(string stamp)
    {
        var parts = stamp.Trim().Split(',');
        if (parts.Length != 2 || !int.TryParse(parts[0], out int id) || !long.TryParse(parts[1], out long ticks))
            return null;
        try
        {
            var process = Process.GetProcessById(id);
            if (!process.HasExited && process.StartTime.ToUniversalTime().Ticks == ticks) return process;
            process.Dispose();
        }
        catch { /* もう動いていない、または見る権限がない */ }
        return null;
    }

    public static bool IsAlive(string stamp)
    {
        using var process = Find(stamp);
        return process is not null;
    }
}

/// <summary>
/// このツールが起動した Office の控え。作業フォルダの中のファイルに書いておく。
/// ツールが異常終了して Office を終了させられなかったとき、次回の起動時にこの控えを見て片付ける。
/// </summary>
internal static class OfficeLedger
{
    private const string FileName = "office.pids";
    private static readonly string[] OfficeProcessNames = ["WINWORD", "EXCEL", "POWERPNT"];
    private static readonly object Gate = new();
    private static string? _file;

    public static void Use(string workDir)
    {
        lock (Gate) _file = Path.Combine(workDir, FileName);
    }

    public static void Record(Process process)
    {
        lock (Gate)
        {
            if (_file is null) return;
            try { File.AppendAllText(_file, ProcessStamp.Format(process) + Environment.NewLine); }
            catch { /* 控えが書けなくても変換は続ける */ }
        }
    }

    /// <summary>
    /// 控えにある Office のうち、今も動いていて画面を持たないものを終了させる。
    /// 画面に出ているものは、利用者が自分のファイルを開いて使っているので触らない。
    /// </summary>
    public static void KillLeftovers(string workDir)
    {
        string file = Path.Combine(workDir, FileName);
        if (!File.Exists(file)) return;
        foreach (var line in File.ReadAllLines(file))
        {
            try
            {
                using var process = ProcessStamp.Find(line);
                if (process is null) continue;
                if (!OfficeProcessNames.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase)) continue;
                if (process.MainWindowHandle != IntPtr.Zero) continue;
                process.Kill(entireProcessTree: true);
            }
            catch { /* 終了できなくても起動は続ける */ }
        }
    }
}
