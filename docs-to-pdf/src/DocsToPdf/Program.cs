using DocsToPdf.Core;
using DocsToPdf.Ui;

namespace DocsToPdf;

/// <summary>
/// 起動時の指定。ファイルやフォルダのパスのほかに、スクリプトから使うための指定を受け付ける。
///   --out 保存先.pdf   画面の操作なしで結合して終了する（成功 0、失敗 1）
///   --page-numbers / --toc / --duplex   ページ番号・目次・両面印刷用の白紙
///   --open             作成後に PDF を開く
///   --pick [範囲]      最初のファイルの「ページを選んで保存」を、そのページを選んだ状態で開く（例: --pick 1-3 10-12）
///   --split            最初のファイルの「分割」を開く
/// </summary>
internal sealed record StartupOptions(
    string[] Paths, string? Output, bool PageNumbers, bool TableOfContents, bool DuplexBlanks, bool Open,
    StartupDialog? Dialog = null)
{
    public bool IsBatch => Output is not null;

    public static StartupOptions Parse(string[] args)
    {
        var paths = new List<string>();
        string? output = null;
        bool numbers = false, toc = false, duplex = false, open = false;
        StartupDialog? dialog = null;

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            switch (arg.ToLowerInvariant())
            {
                case "--out" or "-o" when i + 1 < args.Length:
                    output = Path.GetFullPath(args[++i]);
                    break;
                case "--page-numbers":
                    numbers = true;
                    break;
                case "--toc":
                    toc = true;
                    break;
                case "--duplex":
                    duplex = true;
                    break;
                case "--open":
                    open = true;
                    break;
                case "--split":
                    dialog = new StartupDialog(Split: true, "");
                    break;
                case "--pick":
                    // 続く引数のうち、ページ指定として読めて、ファイルやフォルダではないものを範囲とみなす
                    var ranges = new List<string>();
                    while (i + 1 < args.Length && StartupDialog.IsRange(args[i + 1])) ranges.Add(args[++i]);
                    dialog = new StartupDialog(Split: false, PageRange.Normalize(string.Join(",", ranges)));
                    break;
                case var pick when pick.StartsWith(StartupDialog.PickPrefix):
                    dialog = new StartupDialog(Split: false, PageRange.Normalize(arg[StartupDialog.PickPrefix.Length..]));
                    break;
                default:
                    if (!string.IsNullOrWhiteSpace(arg)) paths.Add(arg);
                    break;
            }
        }
        return new StartupOptions([.. paths], output, numbers, toc, duplex, open, dialog);
    }
}

/// <summary>
/// ファイルを入れたあとに開く画面。Pages は「ページを選んで保存」で最初から選んでおくページ（空なら選ばない）。
/// すでに起動しているウィンドウへは、ファイルのパスと同じ行の並びに 1 行（--pick=範囲 か --split）で送る。
/// </summary>
internal sealed record StartupDialog(bool Split, string Pages)
{
    public const string PickPrefix = "--pick=";
    private const string SplitLine = "--split";

    public static bool IsRange(string arg) =>
        PageRange.Normalize(arg).Length > 0 && PageRange.IsValid(arg) && !File.Exists(arg) && !Directory.Exists(arg);

    public string ToLine() => Split ? SplitLine : PickPrefix + Pages;

    /// <summary>送られてきた行が画面の指定なら読む。ファイルのパス（いつも絶対パス）なら null。</summary>
    public static StartupDialog? FromLine(string line) =>
        line == SplitLine ? new StartupDialog(true, "")
        : line.StartsWith(PickPrefix, StringComparison.Ordinal) ? new StartupDialog(false, PageRange.Normalize(line[PickPrefix.Length..]))
        : null;
}

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var options = StartupOptions.Parse(args);

        // exe のアイコンや「送る」メニューにファイルを落とすと、そのパスが引数で渡ってくる。
        // すでに起動していれば、そのウィンドウのリストへ追加して終わる。
        SingleInstance? instance = null;
        if (!options.IsBatch)
        {
            // 画面の指定は、パスのあとに 1 行で添える（パスはいつも絶対パスなので、行の頭の「--」と見分けられる）
            string[] lines = [.. options.Paths.Select(Path.GetFullPath), .. options.Dialog is { } d ? [d.ToLine()] : Array.Empty<string>()];
            instance = SingleInstance.Acquire(lines);
            if (instance is null) return 0;
        }

        using (instance)
        {
            ApplicationConfiguration.Initialize();
            // 色のモードは、ウィンドウを 1 つも作らないうちに決める必要がある
            Theme.Apply(AppSettings.Load().ColorMode);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (_, e) =>
            {
                if (options.IsBatch)
                {
                    Console.Error.WriteLine(e.Exception);
                    Environment.Exit(1);
                }
                MessageBox.Show(e.Exception.Message, MainForm.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            TempWorkspace.CleanupStale();
            using var form = new MainForm(options);
            instance?.Listen(form.ReceiveLines);
            Application.Run(form);
            return form.ExitCode;
        }
    }
}
