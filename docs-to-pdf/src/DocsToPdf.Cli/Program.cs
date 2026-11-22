using System.CommandLine;
using System.Diagnostics;
using DocsToPdf.Core;

namespace DocsToPdf.Cli;

/// <summary>終了コード。スクリプトから結果を見分けられるよう、意味ごとに分ける。</summary>
internal static class ExitCodes
{
    public const int Success = 0;
    /// <summary>変換・結合・分割・抜き出しに失敗した（ファイルごとに PDF にするときは、1 件でも失敗した）、または上書きを断った。</summary>
    public const int Failure = 1;
    /// <summary>指定が正しくない（ファイルが見つからない、オプションの組み合わせが違うなど）。</summary>
    public const int Usage = 2;
    /// <summary>Ctrl+C で中止した（シェルの慣習に合わせる）。</summary>
    public const int Cancelled = 130;
}

internal static class Program
{
    private static readonly Option<bool> Quiet = new("--quiet", "-q") { Description = "進み具合を表示しない（エラーと結果のパスだけを出す）" };
    private static readonly Option<bool> Force = new("--force", "-f") { Description = "出力先に同じ名前のファイルがあれば上書きする" };
    private static readonly Option<int?> BookmarkDepth = new("--bookmark-depth")
    {
        Description = "元のファイルのしおり（Word の見出しなど）を N 段まで付ける。0 なら付けない（既定はすべて）",
        HelpName = "N",
    };
    private static readonly Option<int> Timeout = new("--timeout")
    {
        Description = "1 ファイルの変換を待つ時間（分）。過ぎると Office を止めて失敗にする",
        DefaultValueFactory = _ => 5,
        HelpName = "分",
    };

    private static async Task<int> Main(string[] args)
    {
        Timeout.Validators.Add(r =>
        {
            if (r.GetValueOrDefault<int>() < 1) r.AddError("--timeout は 1 以上の分で指定してください。");
        });
        BookmarkDepth.Validators.Add(r =>
        {
            if (r.GetValueOrDefault<int?>() < 0) r.AddError("--bookmark-depth は 0 以上で指定してください。");
        });

        var root = BuildRoot();
        root.Subcommands.Add(BuildMerge());
        root.Subcommands.Add(BuildSplit());
        root.Subcommands.Add(BuildExtract());

        // ライブラリが用意する --help / --version の説明を、ほかと同じ日本語にそろえる
        foreach (var option in root.Options)
        {
            if (option is System.CommandLine.Help.HelpOption) option.Description = "使い方を表示する";
            else if (option is VersionOption) option.Description = "バージョンを表示する";
        }

        // 「送る」メニューや exe へのドロップで起動されたときは、この実行のために開いた画面で動いている
        bool ownWindow = OwnsConsoleWindow();

        int code;
        var parsed = root.Parse(args);
        if (parsed.Errors.Count > 0)
        {
            foreach (var error in parsed.Errors) Console.Error.WriteLine($"エラー: {error.Message}");
            Console.Error.WriteLine("使い方は docs2pdf --help で確認できます。");
            code = ExitCodes.Usage;
        }
        else
        {
            // Ctrl+C を押すと中止の合図が届く。Office の後片付けが終わるまで少し待ってから終了させる
            var config = new InvocationConfiguration { ProcessTerminationTimeout = TimeSpan.FromSeconds(15) };
            code = await parsed.InvokeAsync(config);
        }

        if (ownWindow) WaitBeforeClosing(code);
        return code;
    }

    /// <summary>
    /// この実行のために開いた画面か（つながっているプロセスが自分だけ）。
    /// cmd や PowerShell から実行したときは、その画面を使い続けるので false。
    /// </summary>
    private static bool OwnsConsoleWindow()
    {
        if (Console.IsInputRedirected || Console.IsErrorRedirected) return false;
        var ids = new uint[2];
        return NativeMethods.GetConsoleProcessList(ids, (uint)ids.Length) == 1;
    }

    /// <summary>
    /// 開いた画面が結果を見る前に消えないよう、数秒待ってから閉じる。エラーがあったときは、読めるよう長めに待つ。
    /// キーを押せば待たずに閉じる。Ctrl+C で中止したときは、そのまま閉じる。
    /// </summary>
    private static void WaitBeforeClosing(int code)
    {
        if (code == ExitCodes.Cancelled) return;
        bool failed = code != ExitCodes.Success;
        int seconds = failed ? 10 : 3;
        string head = failed ? "エラーがありました。" : "";
        if (failed) Console.Error.WriteLine();

        for (int left = seconds; left > 0; left--)
        {
            Console.Error.Write($"\r{head}{left} 秒後に閉じます   ");
            for (int tick = 0; tick < 10; tick++)
            {
                if (Console.KeyAvailable)
                {
                    Console.ReadKey(intercept: true);
                    return;
                }
                Thread.Sleep(100);
            }
        }
    }

    /// <summary>
    /// サブコマンドなし: ファイルごとに PDF にする（docs2pdf 報告.docx → 報告.pdf）。
    /// 結合は merge、分割は split、ページの抜き出しは extract。
    /// </summary>
    private static RootCommand BuildRoot()
    {
        var inputs = new Argument<string[]>("入力")
        {
            Description = "PDF にするファイルやフォルダ（フォルダはサブフォルダも含む）。PDF はすでに PDF なので飛ばす",
            Arity = ArgumentArity.ZeroOrMore,
        };
        var output = new Option<string?>("--out", "-o") { Description = "書き出すフォルダ（既定は、それぞれのファイルと同じフォルダ）", HelpName = "フォルダ" };
        var numbers = new Option<bool>("--page-numbers") { Description = "各ページの下に「3 / 25」の形でページ番号を入れる" };

        var root = new RootCommand(
            "Word・Excel・PowerPoint・画像を、ファイルごとに PDF にする（同じ名前の .pdf を隣に作る）。\n" +
            "1 つにまとめるときは merge、PDF を分けるときは split、ページを抜き出すときは extract を使う。")
        {
            inputs, output, numbers, BookmarkDepth, Force, Timeout, Quiet,
        };
        root.SetAction((result, ct) => ConvertEachAsync(
            result.GetValue(inputs) ?? [], result.GetValue(output), result.GetValue(numbers), result.GetValue(BookmarkDepth),
            result.GetValue(Force), result.GetValue(Timeout), new Log(result.GetValue(Quiet)), ct));
        return root;
    }

    private static async Task<int> ConvertEachAsync(string[] inputs, string? output, bool pageNumbers, int? outlineDepth, bool force,
        int timeoutMinutes, Log log, CancellationToken ct)
    {
        if (inputs.Length == 0)
        {
            log.Error("PDF にするファイルかフォルダを指定してください。使い方は docs2pdf --help で確認できます。");
            return ExitCodes.Usage;
        }

        var (found, rejected) = InputFiles.Expand(inputs, sortAll: false);
        foreach (var path in rejected) log.Error($"見つからないか、扱えない形式です: {path}");
        if (rejected.Count > 0) return ExitCodes.Usage;

        // PDF はすでに PDF。前に作った PDF がフォルダにあっても、もう一度 PDF にしないよう飛ばす
        static bool IsPdf(string path) => Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        foreach (var input in inputs.Where(i => IsPdf(i) && File.Exists(i)))
            log.Progress($"すでに PDF なので飛ばします: {input}");
        var files = found.Where(f => !IsPdf(f)).ToList();
        if (files.Count == 0)
        {
            log.Error("PDF にするファイルがありません（Word・Excel・PowerPoint・画像が見つかりません）。");
            return ExitCodes.Usage;
        }

        string? folder = output is null ? null : Path.GetFullPath(output);
        var targets = IndividualPdf.FileNames(files, file => folder ?? Path.GetDirectoryName(file)!);
        if (!CanWrite(targets, force, log)) return ExitCodes.Failure;

        return await RunAsync(log, timeoutMinutes, async session =>
        {
            // 変換できたものから順に書き出す（途中で止めても、終わった分は残る）
            int written = 0;
            await session.ConvertEachAsync(files, ct, async (i, pdf) =>
            {
                var part = new MergePart(pdf.PdfPath, Path.GetFileNameWithoutExtension(files[i]),
                    Enumerable.Range(0, pdf.PageCount).ToArray(), pdf.Outline);
                Directory.CreateDirectory(Path.GetDirectoryName(targets[i])!);
                await Task.Run(() => IndividualPdf.Write(part, targets[i], pageNumbers, ct, outlineDepth), ct);
                log.Result(targets[i]);
                written++;
            });

            int failed = files.Count - written;
            if (failed > 0)
            {
                log.Error($"{failed} 件のファイルを PDF にできませんでした（ほかの {written} 件は書き出しました）。");
                return ExitCodes.Failure;
            }
            log.Progress($"{written} 件のファイルを PDF にしました");
            return ExitCodes.Success;
        });
    }

    private static Command BuildMerge()
    {
        var inputs = new Argument<string[]>("入力")
        {
            Description = "結合するファイルやフォルダ（フォルダはサブフォルダも含む）。書いた順に並べ、フォルダの中はファイル名の順",
            Arity = ArgumentArity.OneOrMore,
        };
        var output = new Option<string>("--out", "-o") { Description = "作る PDF のパス", Required = true, HelpName = "出力.pdf" };
        var toc = new Option<bool>("--toc") { Description = "先頭に目次ページを付ける（行を押すとそのページへ飛ぶ）" };
        var numbers = new Option<bool>("--page-numbers") { Description = "各ページの下に「3 / 25」の形で通し番号を入れる" };
        var duplex = new Option<bool>("--duplex") { Description = "各ファイルが奇数ページから始まるよう、白紙を挟む（両面印刷用）" };
        var title = new Option<string?>("--title") { Description = "PDF の文書のタイトル（既定は出力のファイル名）", HelpName = "題名" };
        var open = new Option<bool>("--open") { Description = "作ったあと、既定のアプリで開く" };
        var flat = new Option<bool>("--no-folder-bookmarks")
        {
            Description = "フォルダを渡したとき、しおりをフォルダごとにまとめない（既定では、フォルダの段の下にファイルのしおりを入れる）",
        };

        var command = new Command("merge", "ファイルを PDF にして、ファイルごとのしおりを付けて 1 つに結合する")
        {
            inputs, output, toc, numbers, duplex, title, BookmarkDepth, flat, open, Force, Timeout, Quiet,
        };
        command.SetAction((result, ct) => MergeAsync(
            result.GetValue(inputs)!, result.GetValue(output)!,
            new MergeOptions(result.GetValue(numbers), result.GetValue(toc), result.GetValue(duplex), result.GetValue(title) ?? "")
            {
                OutlineDepth = result.GetValue(BookmarkDepth),
            },
            !result.GetValue(flat), result.GetValue(open), result.GetValue(Force), result.GetValue(Timeout), new Log(result.GetValue(Quiet)), ct));
        return command;
    }

    private static async Task<int> MergeAsync(string[] inputs, string output, MergeOptions options, bool folderBookmarks, bool open, bool force,
        int timeoutMinutes, Log log, CancellationToken ct)
    {
        var expanded = InputFiles.Expand(inputs, sortAll: false); // 書いた順に結合する
        var (files, rejected) = expanded;
        foreach (var path in rejected) log.Error($"見つからないか、扱えない形式です: {path}");
        if (rejected.Count > 0) return ExitCodes.Usage;
        if (files.Count == 0)
        {
            log.Error("結合できるファイルがありません（フォルダの中に対象のファイルが見つかりません）。");
            return ExitCodes.Usage;
        }

        string target = Path.GetFullPath(output);
        if (files.Contains(target, StringComparer.OrdinalIgnoreCase))
        {
            log.Error($"出力先が入力のファイルと同じです: {target}");
            return ExitCodes.Usage;
        }
        if (!CanWrite([target], force, log)) return ExitCodes.Failure;

        return await RunAsync(log, timeoutMinutes, async session =>
        {
            var converted = await session.ConvertAllAsync(files, ct);
            if (converted is null)
            {
                log.Error("結合できないファイルがあるため、PDF は作りませんでした。");
                return ExitCodes.Failure;
            }

            var folders = folderBookmarks ? BookmarkFolders.Of(files, files.Select(expanded.RootOf).ToList()) : null;
            var parts = files.Select((file, i) =>
            {
                var pdf = converted[i];
                return new MergePart(pdf.PdfPath, Path.GetFileNameWithoutExtension(file), Enumerable.Range(0, pdf.PageCount).ToArray(), pdf.Outline)
                {
                    Folders = folders?[i] ?? [],
                };
            }).ToList();
            if (options.DocumentTitle.Length == 0) options = options with { DocumentTitle = Path.GetFileNameWithoutExtension(target) };

            log.Progress("結合しています...");
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            int pages = await Task.Run(() => PdfMerger.Merge(parts, options, target, ct: ct), ct);

            log.Progress($"作成しました（{files.Count} ファイル、{pages} ページ）");
            log.Result(target);
            if (open) ShellOpen(target, log);
            return ExitCodes.Success;
        });
    }

    private static Command BuildSplit()
    {
        var input = new Argument<string>("入力") { Description = "分割するファイル（PDF のほか、Word・Excel・PowerPoint・画像も PDF にしてから分ける）" };
        var output = new Option<string?>("--out", "-o") { Description = "書き出すフォルダ（既定は入力と同じフォルダ）", HelpName = "フォルダ" };
        var every = new Option<int?>("--every") { Description = "N ページごとに分ける", HelpName = "N" };
        var pages = new Option<string[]>("--pages")
        {
            Description = "この範囲を 1 ファイルにする。複数回指定すると、その数だけファイルを作る（例: --pages 1-3 --pages 4-）",
            HelpName = "範囲",
            AllowMultipleArgumentsPerToken = false,
        };
        var bookmarks = new Option<bool>("--bookmarks") { Description = "先頭の階層のしおりごとに分ける（章ごとに分けるとき）" };

        var command = new Command("split", "PDF を複数のファイルに分ける。--every・--pages・--bookmarks のどれか 1 つで分け方を選ぶ")
        {
            input, output, every, pages, bookmarks, Force, Timeout, Quiet,
        };
        command.SetAction((result, ct) => SplitAsync(
            result.GetValue(input)!, result.GetValue(output), result.GetValue(every), result.GetValue(pages) ?? [],
            result.GetValue(bookmarks), result.GetValue(Force), result.GetValue(Timeout), new Log(result.GetValue(Quiet)), ct));
        return command;
    }

    private static async Task<int> SplitAsync(string input, string? output, int? every, string[] ranges, bool byBookmarks,
        bool force, int timeoutMinutes, Log log, CancellationToken ct)
    {
        int modes = (every is null ? 0 : 1) + (ranges.Length == 0 ? 0 : 1) + (byBookmarks ? 1 : 0);
        if (modes != 1)
        {
            log.Error("分け方を --every・--pages・--bookmarks のどれか 1 つで指定してください。");
            return ExitCodes.Usage;
        }
        if (every is < 1)
        {
            log.Error("--every は 1 以上で指定してください。");
            return ExitCodes.Usage;
        }

        string source = Path.GetFullPath(input);
        if (!File.Exists(source) || !FileEntry.TryGetKind(source, out _))
        {
            log.Error($"見つからないか、扱えない形式です: {input}");
            return ExitCodes.Usage;
        }
        string folder = Path.GetFullPath(output ?? Path.GetDirectoryName(source)!);
        string baseName = Path.GetFileNameWithoutExtension(source);

        return await RunAsync(log, timeoutMinutes, async session =>
        {
            var converted = await session.ConvertAllAsync([source], ct);
            if (converted is null) return ExitCodes.Failure;
            var pdf = converted[0];

            IReadOnlyList<PdfSplitter.Piece> pieces;
            try
            {
                pieces = every is int n ? PdfSplitter.Every(pdf.PageCount, n)
                    : ranges.Length > 0 ? PdfSplitter.Ranges(ranges, pdf.PageCount)
                    : PdfSplitter.ByBookmarks(pdf.Outline, pdf.PageCount);
            }
            catch (FormatException ex)
            {
                log.Error(ex.Message);
                return ExitCodes.Usage;
            }
            if (pieces.Count == 0)
            {
                log.Error("この PDF にはしおりがないため、しおりごとには分けられません。");
                return ExitCodes.Failure;
            }

            var paths = PdfSplitter.FileNames(baseName, pieces, folder);
            if (paths.Contains(source, StringComparer.OrdinalIgnoreCase))
            {
                log.Error($"書き出すファイルが入力のファイルと同じ名前になります: {source}");
                return ExitCodes.Usage;
            }
            if (!CanWrite(paths, force, log)) return ExitCodes.Failure;

            log.Progress($"{pieces.Count} ファイルに分けています...");
            Directory.CreateDirectory(folder);
            await Task.Run(() => PdfSplitter.Split(pdf.PdfPath, pdf.Outline, baseName, pieces, paths, ct: ct), ct);
            foreach (var path in paths) log.Result(path);
            log.Progress($"分割しました（全 {pdf.PageCount} ページ → {pieces.Count} ファイル）");
            return ExitCodes.Success;
        });
    }

    private static Command BuildExtract()
    {
        var input = new Argument<string>("入力") { Description = "ページを抜き出すファイル（PDF のほか、Word・Excel・PowerPoint・画像も PDF にしてから抜き出す）" };
        var pages = new Option<string[]>("--pages")
        {
            Description = "抜き出すページ。空白かカンマで区切って並べる（例: --pages 1-3 10-12）。書いた順ではなく、元のページの順に並べる",
            HelpName = "範囲",
            Required = true,
            Arity = ArgumentArity.OneOrMore,
            AllowMultipleArgumentsPerToken = true,
        };
        var output = new Option<string?>("--out", "-o")
        {
            Description = "作る PDF のパス（既定は入力と同じフォルダの「元の名前_1-3,10-12.pdf」）",
            HelpName = "出力.pdf",
        };
        var numbers = new Option<bool>("--page-numbers") { Description = "各ページの下に「3 / 25」の形で、抜き出した中での番号を入れる" };
        var open = new Option<bool>("--open") { Description = "作ったあと、既定のアプリで開く" };

        var command = new Command("extract", "選んだページだけを 1 つの PDF にする（1-3 と 10-12 だけ印刷したいときなど）")
        {
            input, pages, output, numbers, open, Force, Timeout, Quiet,
        };
        command.SetAction((result, ct) => ExtractAsync(
            result.GetValue(input)!, result.GetValue(pages) ?? [], result.GetValue(output), result.GetValue(numbers),
            result.GetValue(open), result.GetValue(Force), result.GetValue(Timeout), new Log(result.GetValue(Quiet)), ct));
        return command;
    }

    private static async Task<int> ExtractAsync(string input, string[] ranges, string? output, bool pageNumbers, bool open,
        bool force, int timeoutMinutes, Log log, CancellationToken ct)
    {
        string range = string.Join(",", ranges.Select(PageRange.Normalize).Where(r => r.Length > 0));
        if (range.Length == 0 || !PageRange.IsValid(range))
        {
            log.Error($"ページ指定「{string.Join(" ", ranges)}」を読めません。「1-3 10-12」「5,8-」のように書いてください。");
            return ExitCodes.Usage;
        }

        string source = Path.GetFullPath(input);
        if (!File.Exists(source) || !FileEntry.TryGetKind(source, out _))
        {
            log.Error($"見つからないか、扱えない形式です: {input}");
            return ExitCodes.Usage;
        }
        string stem = Path.GetFileNameWithoutExtension(source);

        return await RunAsync(log, timeoutMinutes, async session =>
        {
            var converted = await session.ConvertAllAsync([source], ct);
            if (converted is null) return ExitCodes.Failure;
            var pdf = converted[0];

            var pages = PageRange.Resolve(range, pdf.PageCount);
            if (pages.Count == 0)
            {
                log.Error($"ページ指定「{range}」に当てはまるページがありません（全 {pdf.PageCount} ページ）。");
                return ExitCodes.Usage;
            }

            // 名前は画面の「ページを選んで保存」と同じ（元の名前_1-3,10-12.pdf）
            string label = PageRange.Format(pages, pdf.PageCount) is { Length: > 0 } l ? l : "全ページ";
            string target = Path.GetFullPath(output
                ?? Path.Combine(Path.GetDirectoryName(source)!, PdfSplitter.SafeFileName($"{stem}_{label}") + ".pdf"));
            if (string.Equals(target, source, StringComparison.OrdinalIgnoreCase))
            {
                log.Error($"出力先が入力のファイルと同じです: {target}");
                return ExitCodes.Usage;
            }
            if (!CanWrite([target], force, log)) return ExitCodes.Failure;

            log.Progress($"{pages.Count} ページを抜き出しています...");
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            var part = new MergePart(pdf.PdfPath, stem, pages, pdf.Outline);
            int written = await Task.Run(() => IndividualPdf.Write(part, target, pageNumbers, ct), ct);
            log.Progress($"作成しました（全 {pdf.PageCount} ページから {written} ページ）");
            log.Result(target);
            if (open) ShellOpen(target, log);
            return ExitCodes.Success;
        });
    }

    /// <summary>変換の仕組みを用意して処理を行い、中止と予期しない失敗を終了コードに直す。</summary>
    private static async Task<int> RunAsync(Log log, int timeoutMinutes, Func<Session, Task<int>> work)
    {
        try
        {
            using var session = new Session(TimeSpan.FromMinutes(timeoutMinutes), log);
            return await work(session);
        }
        catch (OperationCanceledException)
        {
            log.Error("中止しました。");
            return ExitCodes.Cancelled;
        }
        catch (Exception ex)
        {
            log.Error(ex.Message);
            return ExitCodes.Failure;
        }
    }

    /// <summary>上書きしてよいかを確かめる。--force がなければ、既にあるファイルには書かない。</summary>
    private static bool CanWrite(IReadOnlyList<string> paths, bool force, Log log)
    {
        if (force) return true;
        var existing = paths.Where(File.Exists).ToList();
        if (existing.Count == 0) return true;
        foreach (var path in existing.Take(5)) log.Error($"既にあります: {path}");
        if (existing.Count > 5) log.Error($"ほか {existing.Count - 5} 件");
        log.Error("上書きするときは --force を付けてください。");
        return false;
    }

    private static void ShellOpen(string path, Log log)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })?.Dispose(); }
        catch (Exception ex) { log.Error($"開けませんでした: {ex.Message}"); }
    }
}
