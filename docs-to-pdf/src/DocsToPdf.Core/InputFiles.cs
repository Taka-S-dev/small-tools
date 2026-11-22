namespace DocsToPdf.Core;

/// <summary>渡されたファイルとフォルダから、扱えるファイルを集める。画面とコマンドラインで同じ決まりにする。</summary>
internal static class InputFiles
{
    /// <param name="Files">扱えるファイル（完全なパス、重複なし）。</param>
    /// <param name="Rejected">見つからない、または扱えない形式だったもの（渡されたままの表記）。</param>
    public sealed record Result(IReadOnlyList<string> Files, IReadOnlyList<string> Rejected)
    {
        /// <summary>フォルダの中から見つけたファイル → そのフォルダ（渡されたフォルダ。しおりをフォルダごとにまとめるのに使う）。</summary>
        public IReadOnlyDictionary<string, string> Roots { get; init; } = new Dictionary<string, string>();

        /// <summary>そのファイルを見つけたフォルダ。直接渡されたファイルなら null。</summary>
        public string? RootOf(string file) => Roots.TryGetValue(file, out var root) ? root : null;
    }

    /// <summary>
    /// フォルダはサブフォルダの中も含めて探し、その中はパスの順（数字は数値として）に並べる。
    /// Office が作るロックファイル（~$）は除く。同じファイルが重なれば、最初の 1 つだけを残す。
    /// </summary>
    /// <param name="sortAll">
    /// true なら全体をパスの順に並べ直す（エクスプローラーからのドロップは並びが当てにならないため）。
    /// false なら渡された順を守る（コマンドラインで書いた順が、そのまま結合の順になる）。
    /// </param>
    /// <param name="progress">大きなフォルダを探している間の進み具合（調べたファイルの数、見つけた対象の数）。</param>
    /// <param name="ct">中止の合図。中止されたら OperationCanceledException を投げる。</param>
    public static Result Expand(IEnumerable<string> paths, bool sortAll = true,
        IProgress<(int Scanned, int Found)>? progress = null, CancellationToken ct = default)
    {
        var files = new List<string>();
        var rejected = new List<string>();
        // 同じファイルが重なったときは最初の 1 つを残すので、フォルダも最初に見つけたときのものにする
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var roots = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        int scanned = 0;
        foreach (var raw in paths)
        {
            ct.ThrowIfCancellationRequested();
            string path;
            try { path = Path.GetFullPath(raw); }
            catch
            {
                rejected.Add(raw);
                continue;
            }

            if (Directory.Exists(path))
            {
                var options = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
                var found = new List<string>();
                foreach (var file in Directory.EnumerateFiles(path, "*", options))
                {
                    // ドライブ直下のような大きなフォルダでも止められるよう、1 件ずつ中止の合図を見る
                    ct.ThrowIfCancellationRequested();
                    if (FileEntry.TryGetKind(file, out _)) found.Add(file);
                    if (++scanned % 500 == 0) progress?.Report((scanned, files.Count + found.Count));
                }
                found.Sort(NaturalStringComparer.Instance);
                files.AddRange(found);
                foreach (var file in found)
                    if (seen.Add(file)) roots[file] = path;
            }
            else if (File.Exists(path) && FileEntry.TryGetKind(path, out _))
            {
                files.Add(path);
                seen.Add(path);
            }
            else
            {
                rejected.Add(raw);
            }
        }

        var unique = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (sortAll) unique.Sort(NaturalStringComparer.Instance);
        return new Result(unique, rejected) { Roots = roots };
    }
}
