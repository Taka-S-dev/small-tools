namespace DocsToPdf.Core;

/// <summary>
/// 結合せずに、ファイルごとに 1 つの PDF を書き出す。名前は「元のファイル名.pdf」。
/// </summary>
internal static class IndividualPdf
{
    /// <summary>
    /// 書き出す PDF のパスを決める。folderOf は元のファイルから書き出すフォルダを返す。
    /// 同じフォルダで名前が重なるもの（報告.docx と 報告.xlsx）は「報告 (docx).pdf」のように拡張子を添え、
    /// 入力のファイルそのもの（報告.pdf を同じフォルダに書く場合）と reserved のパスには、決して書かない。
    /// </summary>
    public static IReadOnlyList<string> FileNames(IReadOnlyList<string> sources, Func<string, string> folderOf,
        IEnumerable<string>? reserved = null)
    {
        var inputs = new HashSet<string>(sources.Concat(reserved ?? []).Select(Path.GetFullPath), StringComparer.OrdinalIgnoreCase);
        var folders = sources.Select(s => Path.GetFullPath(folderOf(s))).ToList();
        var stems = sources.Select(s => PdfSplitter.SafeFileName(Path.GetFileNameWithoutExtension(s))).ToList();
        var crowded = Enumerable.Range(0, sources.Count)
            .GroupBy(i => Path.Combine(folders[i], stems[i]), StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .SelectMany(g => g)
            .ToHashSet();

        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paths = new List<string>(sources.Count);
        for (int i = 0; i < sources.Count; i++)
        {
            string stem = stems[i].Length > 0 ? stems[i] : "文書";
            string extension = Path.GetExtension(sources[i]).TrimStart('.').ToLowerInvariant();
            string name = crowded.Contains(i) ? $"{stem} ({extension})" : stem;
            string path = Path.Combine(folders[i], name + ".pdf");
            if (inputs.Contains(path) && !crowded.Contains(i))
            {
                name = $"{stem} ({extension})";
                path = Path.Combine(folders[i], name + ".pdf");
            }
            for (int n = 2; inputs.Contains(path) || !used.Add(path); n++)
                path = Path.Combine(folders[i], $"{name} ({n}).pdf");
            paths.Add(path);
        }
        return paths;
    }

    /// <summary>
    /// 1 ファイル分を書き出す。使うページと回したページを反映し、元のしおりはそのまま一番上の階層に置く。
    /// 目次と両面印刷用の白紙は、1 ファイルでは意味がないので付けない。
    /// </summary>
    /// <returns>書き出した PDF のページ数。</returns>
    public static int Write(MergePart part, string outputPath, bool pageNumbers, CancellationToken ct = default) =>
        PdfMerger.Merge([part], new MergeOptions(pageNumbers, false, false, part.Title) { FileBookmarks = false }, outputPath, ct: ct);
}
