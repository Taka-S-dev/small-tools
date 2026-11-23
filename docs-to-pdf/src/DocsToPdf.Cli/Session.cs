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
    private readonly bool _small;

    /// <param name="small">Word・Excel・PowerPoint・画像を小さく作る（ConvertRequest.Small）。</param>
    public Session(TimeSpan timeoutPerFile, Log log, bool small = false)
    {
        _log = log;
        _small = small;
        TempWorkspace.CleanupStale(); // 前に異常終了したときの残りを片付ける
        _workspace = TempWorkspace.Create();
        _service = new ConversionService(_workspace.FolderPath, timeoutPerFile);
    }

    /// <summary>ファイルを順に PDF にする。1 つでも失敗すれば、理由を出して null を返す。</summary>
    public Task<IReadOnlyList<ConvertedPdf>?> ConvertAllAsync(IReadOnlyList<string> files, CancellationToken ct) =>
        ConvertAllAsync(files.Select(RequestOf).ToList(), ct);

    /// <summary>シートや幅の指定つきの依頼を作る（小さく作るかは、この実行の指定に従う）。</summary>
    public ConvertRequest RequestOf(string file, IReadOnlyList<string>? sheets, bool fitWidth)
    {
        FileEntry.TryGetKind(file, out var kind);
        return new ConvertRequest(file, kind, sheets, fitWidth, null, _small);
    }

    /// <summary>シートや幅の指定つきで、順に PDF にする。1 つでも失敗すれば、理由を出して null を返す。</summary>
    public async Task<IReadOnlyList<ConvertedPdf>?> ConvertAllAsync(IReadOnlyList<ConvertRequest> requests, CancellationToken ct)
    {
        var results = await ConvertEachAsync(requests, ct);
        return results.All(r => r is not null) ? results.Select(r => r!).ToList() : null;
    }

    private ConvertRequest RequestOf(string file) => RequestOf(file, null, false);

    /// <summary>
    /// ファイルを順に PDF にする。失敗したファイルは理由を出し、その位置を null にして返す。
    /// onConverted は 1 件変換できるたびに呼ぶ（ファイルごとに書き出すとき、終わったものから書き出せるように）。
    /// </summary>
    public Task<IReadOnlyList<ConvertedPdf?>> ConvertEachAsync(IReadOnlyList<string> files, CancellationToken ct,
        Func<int, ConvertedPdf, Task>? onConverted = null) =>
        ConvertEachAsync(files.Select(RequestOf).ToList(), ct, onConverted);

    private async Task<IReadOnlyList<ConvertedPdf?>> ConvertEachAsync(IReadOnlyList<ConvertRequest> requests, CancellationToken ct,
        Func<int, ConvertedPdf, Task>? onConverted = null)
    {
        var files = requests.Select(r => r.Path).ToList();
        // 依頼は先にまとめて出す（PDF と画像は Office と並行して進む）。結果は並び順に受け取って表示する
        var tasks = requests.Select(request => _service.ConvertAsync(request, null, ct)).ToList();

        var results = new List<ConvertedPdf?>();
        int failed = 0;
        for (int i = 0; i < files.Count; i++)
        {
            string name = Path.GetFileName(files[i]);
            _log.Status($"{Done(i, files.Count, failed)}　変換中: {name}");
            var outcome = await WaitAsync(tasks[i], ct);
            results.Add(outcome.Pdf);
            if (outcome.Pdf is { } pdf)
            {
                if (onConverted is not null) await onConverted(i, pdf);
            }
            else
            {
                failed++;
                string reason = outcome.NeedsPassword
                    ? "パスワードが必要です（パスワード付きのファイルは画面版でパスワードを入力してください）"
                    : outcome.Error ?? "変換できませんでした";
                _log.Error($"{name}: {reason}");
            }
        }
        _log.Status(Done(files.Count, files.Count, failed));
        _log.EndStatus();
        return results;
    }

    private static string Done(int done, int total, int failed) =>
        failed > 0 ? $"{done}/{total} 件完了（うち失敗 {failed} 件）" : $"{done}/{total} 件完了";

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
        // 起動した Office の終了には数秒かかるので、その間も何をしているかを見せる
        _log.Status("Office を終了しています...");
        _service.Dispose();
        _workspace.Dispose();
        _log.DropStatus();
    }
}

/// <summary>
/// 表示の出し分け。進み具合とエラーは標準エラーに、できたファイルのパスだけを標準出力に出す
/// （スクリプトで出力先を受け取れるようにするため）。
/// 画面に出しているときは、進み具合を 1 行で書き換えて見せ、エラーと結果はその上に残す。
/// 出力をファイルやほかのコマンドへ流しているときは、書き換えずに 1 行ずつ出す（ログが崩れないように）。
/// </summary>
internal sealed class Log(bool quiet)
{
    private static readonly bool Live = !Console.IsErrorRedirected;
    private string? _status;

    public void Progress(string text)
    {
        if (quiet) return;
        Above(() => Console.Error.WriteLine(text));
    }

    public void Error(string text) => Above(() => Console.Error.WriteLine($"エラー: {text}"));

    public void Result(string path) => Above(() => Console.Out.WriteLine(path));

    /// <summary>書き換える 1 行（「2/5 件完了　変換中: …」）。</summary>
    public void Status(string text)
    {
        if (quiet) return;
        if (!Live)
        {
            Console.Error.WriteLine(text);
            return;
        }
        _status = text;
        Draw();
    }

    /// <summary>書き換えていた行を、最後の内容のまま確定する。</summary>
    public void EndStatus()
    {
        if (_status is null) return;
        _status = null;
        Console.Error.WriteLine();
    }

    /// <summary>書き換えていた行を消す（一時的な表示だったとき）。</summary>
    public void DropStatus()
    {
        if (_status is null) return;
        _status = null;
        Clear();
    }

    /// <summary>書き換えている行をいったん消して text を書き、行を描き直す。</summary>
    private void Above(Action write)
    {
        if (_status is null)
        {
            write();
            return;
        }
        Clear();
        write();
        Draw();
    }

    /// <summary>行の頭に戻って書き、前の内容が長かったときの残りを空白で消す。折り返さないよう幅より 1 マス短くする。</summary>
    private void Draw()
    {
        int cells = Width() - 1;
        string line = Fit(_status!, cells);
        Console.Error.Write("\r" + line + new string(' ', Math.Max(0, cells - Cells(line))) + "\r" + line);
    }

    private static void Clear() => Console.Error.Write("\r" + new string(' ', Width() - 1) + "\r");

    private static int Width()
    {
        try { return Math.Max(20, Console.WindowWidth); }
        catch (IOException) { return 80; }
    }

    /// <summary>画面の幅（全角は 2 マス）に収まるよう、末尾を「…」で切る。</summary>
    internal static string Fit(string text, int cells)
    {
        if (Cells(text) <= cells) return text;
        var sb = new System.Text.StringBuilder();
        int used = 0;
        foreach (char ch in text)
        {
            int w = IsWide(ch) ? 2 : 1;
            if (used + w > cells - 1) break;
            sb.Append(ch);
            used += w;
        }
        return sb.Append('…').ToString();
    }

    internal static int Cells(string text) => text.Sum(ch => IsWide(ch) ? 2 : 1);

    /// <summary>コンソールで 2 マスを使う文字（日本語の文字、全角の記号など）。</summary>
    private static bool IsWide(char ch) =>
        ch is >= 'ᄀ' and <= 'ᅟ' or >= '⺀' and <= '꓏' or >= '가' and <= '힣'
            or >= '豈' and <= '﫿' or >= '︰' and <= '﹏' or >= '＀' and <= '｠' or >= '￠' and <= '￦';
}
