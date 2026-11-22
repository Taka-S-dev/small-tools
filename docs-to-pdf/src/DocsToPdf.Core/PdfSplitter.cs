using System.Text;

namespace DocsToPdf.Core;

/// <summary>
/// PDF を複数のファイルに分ける。分け方は「N ページごと」「ページ範囲ごと」「先頭の階層のしおりごと」。
/// 書き出しは結合と同じ仕組み（PdfMerger）を使い、元のしおりのうち、その範囲に入るものをそのまま一番上の階層に残す。
/// </summary>
internal static class PdfSplitter
{
    /// <summary>1 つの出力ファイル。Pages は元の PDF の 0 始まりのページ（ページ順）。Label はファイル名に添える言葉。</summary>
    public sealed record Piece(IReadOnlyList<int> Pages, string? Label);

    /// <summary>N ページごとに分ける。最後のファイルは N ページに満たないことがある。</summary>
    public static IReadOnlyList<Piece> Every(int pageCount, int pagesPerFile)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(pagesPerFile, 1);
        var pieces = new List<Piece>();
        for (int start = 0; start < pageCount; start += pagesPerFile)
            pieces.Add(new Piece(Enumerable.Range(start, Math.Min(pagesPerFile, pageCount - start)).ToArray(), null));
        return pieces;
    }

    /// <summary>
    /// 範囲ごとに 1 ファイルにする。範囲は「1-3」「5,8-」のようなページ指定（PageRange と同じ書き方）。
    /// 正しく読めない指定や、ページが 1 枚も入らない指定は FormatException にする（黙って全ページにしない）。
    /// </summary>
    public static IReadOnlyList<Piece> Ranges(IEnumerable<string> ranges, int pageCount)
    {
        if (!ranges.Any()) throw new FormatException("分ける範囲を 1 つ以上書いてください（例: 1-3; 4-8; 9-）。");
        var pieces = new List<Piece>();
        foreach (var range in ranges)
        {
            string normalized = PageRange.Normalize(range);
            if (normalized.Length == 0 || !PageRange.IsValid(normalized))
                throw new FormatException($"ページ指定「{range}」を読めません。「1-3」「5,8-」のように書いてください。");
            var pages = PageRange.Resolve(normalized, pageCount);
            if (pages.Count == 0)
                throw new FormatException($"ページ指定「{range}」に当てはまるページがありません（全 {pageCount} ページ）。");
            pieces.Add(new Piece(pages, normalized));
        }
        return pieces;
    }

    /// <summary>
    /// 区切り（新しいファイルが始まるページ、1 始まり）から、すべてのページを順に分ける範囲の書き方を作る。
    /// 先頭のページはいつも 1 つ目のファイルの始まり。例: 12 ページで 5 と 9 から区切る → "1-4; 5-8; 9-12"。
    /// </summary>
    public static string RangesFromStarts(IEnumerable<int> starts, int pageCount)
    {
        if (pageCount <= 0) return "";
        var cuts = starts.Where(s => s >= 1 && s <= pageCount).Append(1).Distinct().Order().ToList();
        var parts = new List<string>(cuts.Count);
        for (int i = 0; i < cuts.Count; i++)
        {
            int from = cuts[i], to = i + 1 < cuts.Count ? cuts[i + 1] - 1 : pageCount;
            parts.Add(from == to ? $"{from}" : $"{from}-{to}");
        }
        return string.Join("; ", parts);
    }

    /// <summary>「1-3; 4-8; 9-」のように「;」か改行で区切った書き方を、範囲ごとに分ける（画面の入力欄用）。</summary>
    public static IReadOnlyList<string> SplitRangeList(string text) =>
        text.Split([';', '；', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>
    /// 先頭の階層のしおりごとに分ける（章ごとに分けるときに使う）。しおりの名前をファイル名に添える。
    /// 最初のしおりより前にページがあれば、それも 1 ファイルにする。しおりが無ければ空を返す。
    /// </summary>
    public static IReadOnlyList<Piece> ByBookmarks(IReadOnlyList<OutlineNode> outline, int pageCount)
    {
        // 同じページを指すしおりや、前へ戻るしおりは、区切りにならないので飛ばす
        var starts = new List<(int Page, string Title)>();
        foreach (var node in outline)
        {
            if (node.PageIndex < 0 || node.PageIndex >= pageCount) continue;
            if (starts.Count > 0 && node.PageIndex <= starts[^1].Page) continue;
            starts.Add((node.PageIndex, node.Title));
        }
        if (starts.Count == 0) return [];

        var pieces = new List<Piece>();
        if (starts[0].Page > 0)
            pieces.Add(new Piece(Enumerable.Range(0, starts[0].Page).ToArray(), "先頭"));
        for (int i = 0; i < starts.Count; i++)
        {
            int end = i + 1 < starts.Count ? starts[i + 1].Page : pageCount;
            pieces.Add(new Piece(Enumerable.Range(starts[i].Page, end - starts[i].Page).ToArray(), starts[i].Title));
        }
        return pieces;
    }

    /// <summary>
    /// 出力するファイルのパスを決める。「元の名前_01.pdf」の形にし、範囲やしおりの名前があれば添える。
    /// 番号の桁はファイルの数に合わせ、エクスプローラーで並べたときに順番どおりになるようにする。
    /// </summary>
    public static IReadOnlyList<string> FileNames(string baseName, IReadOnlyList<Piece> pieces, string folder)
    {
        int digits = Math.Max(2, pieces.Count.ToString().Length);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var names = new List<string>();
        for (int i = 0; i < pieces.Count; i++)
        {
            string number = (i + 1).ToString().PadLeft(digits, '0');
            string label = pieces[i].Label is { Length: > 0 } l ? "_" + SafeFileName(l) : "";
            string name = $"{baseName}_{number}{label}";
            if (name.Length > 120) name = name[..120];
            string path = Path.Combine(folder, name + ".pdf");
            for (int n = 2; !used.Add(path); n++)
                path = Path.Combine(folder, $"{name}_{n}.pdf");
            names.Add(path);
        }
        return names;
    }

    /// <summary>ファイル名に使えない文字を「_」に置き換える。</summary>
    public static string SafeFileName(string text)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(text.Length);
        foreach (char ch in text.Trim())
            sb.Append(invalid.Contains(ch) || char.IsControl(ch) ? '_' : ch);
        return sb.ToString().TrimEnd('.', ' ');
    }

    // 画面では、結合から外したページを除いた「見えている順の 1, 2, 3...」で数える。
    // そのため、分け方は「使うページの並び（位置）」で決め、書き出す前に元のページ番号へ戻す。

    /// <summary>しおりの行き先を、元のページ番号から「使うページの中での位置」に置き換える。使わないページを指すしおりは除く（子は残せるものだけ同じ階層に上げる）。</summary>
    public static IReadOnlyList<OutlineNode> Restrict(IReadOnlyList<OutlineNode> outline, IReadOnlyList<int> used)
    {
        var position = new Dictionary<int, int>();
        for (int i = 0; i < used.Count; i++) position.TryAdd(used[i], i);

        List<OutlineNode> Walk(IReadOnlyList<OutlineNode> nodes)
        {
            var result = new List<OutlineNode>();
            foreach (var node in nodes)
            {
                var children = Walk(node.Children);
                if (position.TryGetValue(node.PageIndex, out int at)) result.Add(new OutlineNode(node.Title, at, children));
                else result.AddRange(children);
            }
            return result;
        }
        return Walk(outline);
    }

    /// <summary>位置で決めた分け方を、元の PDF のページ番号に戻す。</summary>
    public static IReadOnlyList<Piece> ToSource(IReadOnlyList<Piece> pieces, IReadOnlyList<int> used) =>
        pieces.Select(p => p with { Pages = p.Pages.Select(i => used[i]).ToArray() }).ToList();

    /// <summary>分けて書き出す。書き出したファイルのパスを順に返す。</summary>
    /// <param name="title">各ファイルの先頭のしおりと文書の題名に使う名前（元のファイル名など）。</param>
    /// <param name="rotations">さらに右へ回すページ（元のページ番号 → 90 / 180 / 270）。</param>
    public static IReadOnlyList<string> Split(string pdfPath, IReadOnlyList<OutlineNode> outline, string title,
        IReadOnlyList<Piece> pieces, IReadOnlyList<string> outputPaths, IProgress<int>? progress = null, CancellationToken ct = default,
        IReadOnlyDictionary<int, int>? rotations = null)
    {
        if (pieces.Count != outputPaths.Count) throw new ArgumentException("ファイルの数と出力先の数が違います。", nameof(outputPaths));
        for (int i = 0; i < pieces.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            string pieceTitle = pieces[i].Label is { Length: > 0 } label ? $"{title} - {label}" : $"{title} ({i + 1})";
            var part = new MergePart(pdfPath, pieceTitle, pieces[i].Pages, outline, rotations);
            PdfMerger.Merge([part], MergeOptions.None with { DocumentTitle = pieceTitle, FileBookmarks = false }, outputPaths[i], ct: ct);
            progress?.Report(i + 1);
        }
        return outputPaths;
    }
}
