using DocsToPdf.Core;

namespace DocsToPdf.Cli;

/// <summary>
/// 1 回の実行で使う作業フォルダと変換の仕組み。Dispose で、起動した Office を終了させて作業フォルダを消す。
/// 中止（Ctrl+C）されたときも Dispose が走るので、Office が残らない。
/// </summary>
internal sealed class Session : IDisposable
{
    private readonly TempWorkspace _workspace;
    private readonly ConversionService _service;
    private readonly Log _log;

    public Session(TimeSpan timeoutPerFile, Log log)
    {
        _log = log;
        TempWorkspace.CleanupStale(); // 前に異常終了したときの残りを片付ける
        _workspace = TempWorkspace.Create();
        _service = new ConversionService(_workspace.FolderPath, timeoutPerFile);
    }

    /// <summary>ファイルを順に PDF にする。1 つでも失敗すれば、理由を出して null を返す。</summary>
    public async Task<IReadOnlyList<ConvertedPdf>?> ConvertAllAsync(IReadOnlyList<string> files, CancellationToken ct)
    {
        // 依頼は先にまとめて出す（PDF と画像は Office と並行して進む）。結果は並び順に受け取って表示する
        var tasks = files.Select(file =>
        {
            FileEntry.TryGetKind(file, out var kind);
            return _service.ConvertAsync(new ConvertRequest(file, kind, null, false, null), null, ct);
        }).ToList();

        var results = new List<ConvertedPdf>();
        int failed = 0;
        for (int i = 0; i < files.Count; i++)
        {
            string name = Path.GetFileName(files[i]);
            _log.Progress($"[{i + 1}/{files.Count}] {name}");
            var outcome = await WaitAsync(tasks[i], ct);
            if (outcome.Pdf is { } pdf)
            {
                results.Add(pdf);
                _log.Progress($"      {pdf.PageCount} ページ");
            }
            else
            {
                failed++;
                string reason = outcome.NeedsPassword
                    ? "パスワードが必要です（パスワード付きのファイルは画面版で入力して結合してください）"
                    : outcome.Error ?? "変換できませんでした";
                _log.Error($"{name}: {reason}");
            }
        }
        return failed == 0 ? results : null;
    }

    /// <summary>
    /// 変換の終わりを待つ。始まった Office の変換は途中で止められないので、
    /// 中止されたらすぐ戻り、あとは Dispose で Office ごと終わらせる。
    /// </summary>
    private static async Task<T> WaitAsync<T>(Task<T> task, CancellationToken ct)
    {
        var cancelled = Task.Delay(Timeout.Infinite, ct);
        if (await Task.WhenAny(task, cancelled) == cancelled) ct.ThrowIfCancellationRequested();
        return await task;
    }

    public void Dispose()
    {
        _service.Dispose();
        _workspace.Dispose();
    }
}

/// <summary>
/// 表示の出し分け。進み具合とエラーは標準エラーに、できたファイルのパスだけを標準出力に出す
/// （スクリプトで出力先を受け取れるようにするため）。
/// </summary>
internal sealed class Log(bool quiet)
{
    public void Progress(string text)
    {
        if (!quiet) Console.Error.WriteLine(text);
    }

    public void Error(string text) => Console.Error.WriteLine($"エラー: {text}");

    public static void Result(string path) => Console.Out.WriteLine(path);
}
