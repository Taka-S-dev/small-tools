using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class PdfSplitterTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    private static int[][] PagesOf(IReadOnlyList<PdfSplitter.Piece> pieces) => pieces.Select(p => p.Pages.ToArray()).ToArray();

    [Theory]
    [InlineData(new int[0], 12, "1-12")]
    [InlineData(new[] { 9, 5 }, 12, "1-4; 5-8; 9-12")]
    [InlineData(new[] { 1, 3 }, 3, "1-2; 3")]            // 先頭は区切らなくても 1 つ目の始まり。1 ページだけのファイルは「3」
    [InlineData(new[] { 2, 2, 0, 13 }, 12, "1; 2-12")]   // 重なりと範囲外は無視する
    public void Ranges_from_cut_points_cover_every_page_in_order(int[] starts, int pages, string expected) =>
        Assert.Equal(expected, PdfSplitter.RangesFromStarts(starts, pages));

    [Fact]
    public void Ranges_made_from_cut_points_split_where_the_cuts_are()
    {
        string text = PdfSplitter.RangesFromStarts([4, 7], 9);

        var pieces = PdfSplitter.Ranges(PdfSplitter.SplitRangeList(text), 9);

        Assert.Equal([[0, 1, 2], [3, 4, 5], [6, 7, 8]], PagesOf(pieces));
    }

    [Fact]
    public void Every_cuts_into_equal_pieces_and_keeps_the_remainder()
    {
        int[][] expected = [[0, 1], [2, 3], [4]];
        Assert.Equal(expected, PagesOf(PdfSplitter.Every(5, 2)));
    }

    [Fact]
    public void Every_rejects_zero() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfSplitter.Every(5, 0));

    [Fact]
    public void Ranges_makes_one_file_per_range()
    {
        var pieces = PdfSplitter.Ranges(["1-2", "４－"], 5);
        int[][] expected = [[0, 1], [3, 4]];
        Assert.Equal(expected, PagesOf(pieces));
        Assert.Equal(new[] { "1-2", "4-" }, pieces.Select(p => p.Label));
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("3-1")]
    [InlineData("")]
    [InlineData("9-")]
    public void Ranges_refuses_unreadable_or_empty_ranges(string range) =>
        Assert.Throws<FormatException>(() => PdfSplitter.Ranges([range], 5));

    [Fact]
    public void ByBookmarks_starts_a_file_at_each_top_level_bookmark()
    {
        OutlineNode Node(string title, int page) => new(title, page, []);
        var outline = new[] { Node("序章", 1), Node("第一章", 2), Node("同じページ", 2), Node("第二章", 4) };

        var pieces = PdfSplitter.ByBookmarks(outline, 6);

        int[][] expected = [[0], [1], [2, 3], [4, 5]];
        Assert.Equal(expected, PagesOf(pieces));
        Assert.Equal(new[] { "先頭", "序章", "第一章", "第二章" }, pieces.Select(p => p.Label));
    }

    [Fact]
    public void ByBookmarks_returns_nothing_without_bookmarks() =>
        Assert.Empty(PdfSplitter.ByBookmarks([], 3));

    [Fact]
    public void FileNames_number_in_order_and_make_labels_safe()
    {
        var pieces = Enumerable.Range(0, 12).Select(i => new PdfSplitter.Piece([i], i == 0 ? "A/B:C?" : i == 1 ? "x" : null)).ToList();
        pieces[2] = pieces[1];

        var names = PdfSplitter.FileNames("報告", pieces, @"C:\out").Select(Path.GetFileName).ToList();

        Assert.Equal("報告_01_A_B_C_.pdf", names[0]);
        Assert.Equal("報告_02_x.pdf", names[1]);
        Assert.Equal("報告_12.pdf", names[11]);
        Assert.Equal(names.Count, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void SplitRangeList_accepts_semicolons_and_new_lines() =>
        Assert.Equal(new[] { "1-3", "4,6", "9-" }, PdfSplitter.SplitRangeList(" 1-3; 4,6 ；\r\n9- ;"));

    [Fact]
    public void Ranges_requires_at_least_one_range() =>
        Assert.Throws<FormatException>(() => PdfSplitter.Ranges([], 5));

    [Fact]
    public void Restrict_renumbers_bookmarks_to_the_pages_in_use_and_lifts_orphaned_children()
    {
        // 元は 6 ページ。2 ページ目（index 1）を外して使う
        int[] used = [0, 2, 3, 4, 5];
        var outline = new[]
        {
            new OutlineNode("一", 0, []),
            new OutlineNode("外したページ", 1, [new OutlineNode("その中の節", 3, [])]),
            new OutlineNode("五", 4, []),
        };

        var restricted = PdfSplitter.Restrict(outline, used);

        Assert.Equal(new[] { "一", "その中の節", "五" }, restricted.Select(n => n.Title));
        Assert.Equal(new[] { 0, 2, 3 }, restricted.Select(n => n.PageIndex));
    }

    [Fact]
    public void ToSource_maps_positions_back_to_original_page_numbers()
    {
        int[] used = [0, 2, 3, 4, 5];
        var pieces = PdfSplitter.ToSource(PdfSplitter.Every(used.Length, 2), used);
        int[][] expected = [[0, 2], [3, 4], [5]];
        Assert.Equal(expected, PagesOf(pieces));
    }

    [Fact]
    public void Split_applies_rotations()
    {
        string source = _folder.Pdf("source.pdf", 2);
        var pieces = PdfSplitter.Every(2, 1);
        var paths = PdfSplitter.FileNames("source", pieces, _folder.Path);

        PdfSplitter.Split(source, [], "source", pieces, paths, ct: TestContext.Current.CancellationToken,
            rotations: new Dictionary<int, int> { [1] = 90 });

        Assert.True(PdfInfo.Read(paths[0]).Sizes[0].Width < PdfInfo.Read(paths[0]).Sizes[0].Height);
        Assert.True(PdfInfo.Read(paths[1]).Sizes[0].Width > PdfInfo.Read(paths[1]).Sizes[0].Height, "回したページが横向きになっていない");
    }

    [Fact]
    public void Split_writes_each_piece_and_keeps_bookmarks_inside_it()
    {
        string source = _folder.Pdf("source.pdf", 5, ("一", 0), ("三", 2), ("五", 4));
        var (_, outline) = PdfInfo.Read(source);
        var pieces = PdfSplitter.Every(5, 2);
        var paths = PdfSplitter.FileNames("source", pieces, _folder.Path);

        PdfSplitter.Split(source, outline, "source", pieces, paths, ct: TestContext.Current.CancellationToken);

        Assert.Equal(new[] { 2, 2, 1 }, paths.Select(p => PdfInfo.Read(p).Sizes.Count));
        // 元のしおりを、ファイル名のしおりで包まずにそのまま残す
        var second = PdfInfo.Read(paths[1]).Outline.Single();
        Assert.Equal("三", second.Title);
        Assert.Equal(0, second.PageIndex);
    }
}

public sealed class InputFilesTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    [Fact]
    public void Expand_collects_supported_files_in_natural_order()
    {
        Directory.CreateDirectory(_folder.File("sub"));
        foreach (var name in new[] { "資料10.pdf", "資料2.pdf", @"sub\表.xlsx", "~$開いている.docx", "メモ.txt" })
            File.WriteAllText(_folder.File(name), "");
        string missing = _folder.File("ない.pdf");

        var (files, rejected) = InputFiles.Expand([_folder.Path, _folder.File("資料2.pdf"), missing], ct: TestContext.Current.CancellationToken);

        Assert.Equal(new[] { "sub\\表.xlsx", "資料2.pdf", "資料10.pdf" }, files.Select(f => Path.GetRelativePath(_folder.Path, f)));
        Assert.Equal(new[] { missing }, rejected);
    }

    [Fact]
    public void Expand_stops_when_cancelled()
    {
        File.WriteAllText(_folder.File("a.pdf"), "");
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => InputFiles.Expand([_folder.Path], ct: cancel.Token));
    }

    [Fact]
    public void Expand_reports_progress_while_searching_a_large_folder()
    {
        for (int i = 0; i < 1_200; i++) File.WriteAllText(_folder.File($"f{i}.{(i % 2 == 0 ? "pdf" : "txt")}"), "");
        var reports = new List<(int Scanned, int Found)>();
        var progress = new SynchronousProgress(reports.Add);

        var (files, _) = InputFiles.Expand([_folder.Path], progress: progress, ct: TestContext.Current.CancellationToken);

        Assert.Equal(600, files.Count);
        Assert.Equal(new[] { 500, 1000 }, reports.Select(r => r.Scanned));
        Assert.All(reports, r => Assert.InRange(r.Found, r.Scanned / 2 - 1, r.Scanned / 2 + 1));
    }

    /// <summary>Progress&lt;T&gt; は別のスレッドで知らせるので、テストではその場で受け取る。</summary>
    private sealed class SynchronousProgress(Action<(int, int)> onReport) : IProgress<(int Scanned, int Found)>
    {
        public void Report((int Scanned, int Found) value) => onReport(value);
    }

    [Fact]
    public void Expand_can_keep_the_order_the_paths_were_given_in()
    {
        Directory.CreateDirectory(_folder.File("章"));
        foreach (var name in new[] { "表紙.docx", "末尾.pdf", @"章\2.pdf", @"章\10.pdf" })
            File.WriteAllText(_folder.File(name), "");

        var (files, _) = InputFiles.Expand([_folder.File("末尾.pdf"), _folder.File("章"), _folder.File("表紙.docx"), _folder.File("末尾.pdf")], sortAll: false, ct: TestContext.Current.CancellationToken);

        // 渡した順を守り、フォルダの中だけ数字の順。重なった 2 回目の「末尾」は除く
        Assert.Equal(new[] { "末尾.pdf", "章\\2.pdf", "章\\10.pdf", "表紙.docx" }, files.Select(f => Path.GetRelativePath(_folder.Path, f)));
    }
}
