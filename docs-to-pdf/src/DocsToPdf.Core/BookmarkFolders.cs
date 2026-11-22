namespace DocsToPdf.Core;

/// <summary>しおりのフォルダの段 1 つ。Key はそのフォルダの完全なパスで、名前を変えても段を見分けられるようにする。</summary>
internal sealed record BookmarkFolder(string Key, string Name);

/// <summary>
/// フォルダを入れたときに、しおりをフォルダごとにまとめるための、ファイルごとのフォルダの並び。
/// 入れたフォルダが 1 つだけなら、その名前は段にしない（全部を包むだけで、見分ける役に立たない）。
/// 入れたフォルダが 2 つ以上なら、それぞれの名前を一番上の段にする。直接入れたファイルは、どの段にも入れない。
/// </summary>
internal static class BookmarkFolders
{
    /// <param name="files">ファイル（完全なパス）。</param>
    /// <param name="roots">それぞれのファイルを見つけたフォルダ。直接入れたファイルなら null。</param>
    /// <param name="allRoots">
    /// 入れたフォルダが 1 つだけかを決めるときに数えるフォルダ。files がまだ変換中のものを除いた一部のときに、
    /// 一覧全体のフォルダを渡す（変換が終わるたびに一番上の段が出たり消えたりしないように）。null なら roots。
    /// </param>
    /// <returns>それぞれのファイルを入れる段（上から順）。段に入れないファイルは空。</returns>
    public static IReadOnlyList<IReadOnlyList<BookmarkFolder>> Of(IReadOnlyList<string> files, IReadOnlyList<string?> roots,
        IEnumerable<string?>? allRoots = null)
    {
        var trimmed = roots.Select(Trim).ToList();
        bool single = (allRoots ?? roots).Select(Trim).Where(r => r is not null).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1;

        var result = new List<IReadOnlyList<BookmarkFolder>>(files.Count);
        for (int i = 0; i < files.Count; i++)
        {
            if (trimmed[i] is not { } root)
            {
                result.Add([]);
                continue;
            }
            string relative = Path.GetRelativePath(root, Path.GetDirectoryName(Path.GetFullPath(files[i]))!);
            // フォルダの外（".." で始まる、または別のドライブ）なら段に入れない。"..old" のような名前のフォルダは中として扱う
            bool outside = relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                || Path.IsPathRooted(relative);
            var names = relative == "." || outside
                ? []
                : relative.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries);

            var folders = new List<BookmarkFolder>();
            if (!single) folders.Add(new BookmarkFolder(root, NameOf(root)));
            string key = root;
            foreach (string segment in names)
            {
                key = Path.Combine(key, segment);
                folders.Add(new BookmarkFolder(key, segment));
            }
            result.Add(folders);
        }
        return result;
    }

    /// <summary>フォルダの Key（完全なパス）から付ける名前。ドライブ直下（C:\）のように名前がなければ、パスをそのまま名前にする。</summary>
    public static string NameOf(string key) => Path.GetFileName(key) is { Length: > 0 } name ? name : key;

    private static string? Trim(string? root) => root is null ? null : Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));
}
