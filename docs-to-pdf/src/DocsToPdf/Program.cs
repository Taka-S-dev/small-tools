using DocsToPdf.Core;
using DocsToPdf.Ui;

namespace DocsToPdf;

/// <summary>
/// 起動時の指定。ファイルやフォルダのパスのほかに、スクリプトから使うための指定を受け付ける。
///   --out 保存先.pdf   画面の操作なしで結合して終了する（成功 0、失敗 1）
///   --page-numbers / --toc / --duplex   ページ番号・目次・両面印刷用の白紙
///   --open             作成後に PDF を開く
/// </summary>
internal sealed record StartupOptions(
    string[] Paths, string? Output, bool PageNumbers, bool TableOfContents, bool DuplexBlanks, bool Open)
{
    public bool IsBatch => Output is not null;

    public static StartupOptions Parse(string[] args)
    {
        var paths = new List<string>();
        string? output = null;
        bool numbers = false, toc = false, duplex = false, open = false;

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
                default:
                    if (!string.IsNullOrWhiteSpace(arg)) paths.Add(arg);
                    break;
            }
        }
        return new StartupOptions([.. paths], output, numbers, toc, duplex, open);
    }
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
            instance = SingleInstance.Acquire(options.Paths);
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
            instance?.Listen(form.ReceivePaths);
            Application.Run(form);
            return form.ExitCode;
        }
    }
}
