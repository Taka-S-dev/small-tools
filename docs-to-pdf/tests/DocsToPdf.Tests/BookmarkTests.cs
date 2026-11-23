using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class BookmarkTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    private static CancellationToken Ct => TestContext.Current.CancellationToken;

    private static OutlineNode Node(string title, int page, params OutlineNode[] children) => new(title, page, children);

    private static MergePart Part(string title, int pages, IReadOnlyList<OutlineNode>? outline = null, params string[] folders) =>
        new("unused.pdf", title, Enumerable.Range(0, pages).ToArray(), outline ?? []) { Folders = In(folders) };

    /// <summary>名前の並びから、フォルダの段を作る（Key は名前を / でつないだもの）。</summary>
    private static BookmarkFolder[] In(params string[] names) =>
        names.Select((name, d) => new BookmarkFolder(string.Join("/", names[..(d + 1)]), name)).ToArray();

    /// <summary>しおりを「名前@ページ」の入れ子の文字列にする（比べやすくするため）。</summary>
    private static string Show(IEnumerable<Bookmark> nodes) =>
        string.Join(" ", nodes.Select(n => n.Children.Count == 0 ? $"{n.Title}@{n.Page}" : $"{n.Title}@{n.Page}[{Show(n.Children)}]"));

    private static string Show(IEnumerable<OutlineNode> nodes) =>
        string.Join(" ", nodes.Select(n => n.Children.Count == 0 ? $"{n.Title}@{n.PageIndex}" : $"{n.Title}@{n.PageIndex}[{Show(n.Children)}]"));

    // 第1章(0) ─ 1.1(1) ─ 1.1.1(1)、第2章(3)
    private static readonly OutlineNode[] Chapters =
        [Node("第1章", 0, Node("1.1", 1, Node("1.1.1", 1))), Node("第2章", 3)];

    [Theory]
    [InlineData(null, "資料@0[第1章@0[1.1@1[1.1.1@1]] 第2章@3]")]
    [InlineData(2, "資料@0[第1章@0[1.1@1] 第2章@3]")]
    [InlineData(1, "資料@0[第1章@0 第2章@3]")]
    [InlineData(0, "資料@0")]
    public void Bookmarks_keep_the_original_bookmarks_down_to_the_chosen_depth(int? depth, string expected)
    {
        var bookmarks = PdfMerger.Bookmarks([Part("資料", 4, Chapters)], MergeOptions.None with { OutlineDepth = depth });
        Assert.Equal(expected, Show(bookmarks));
    }

    [Theory]
    [InlineData(null, "資料@0[1.1@0[1.1.1@0] 第2章@2]")]
    [InlineData(1, "資料@0[第2章@2]")]
    public void Bookmarks_lift_children_of_an_excluded_page_only_within_the_depth(int? depth, string expected)
    {
        // 1 ページ目を外すと「第1章」は付かず、その下の「1.1」が 1 段上がる。段の数え方は元のしおりの段のまま
        var part = new MergePart("unused.pdf", "資料", [1, 2, 3], Chapters);
        var bookmarks = PdfMerger.Bookmarks([part], MergeOptions.None with { OutlineDepth = depth });
        Assert.Equal(expected, Show(bookmarks));
    }

    [Fact]
    public void Bookmarks_group_consecutive_files_under_their_folders()
    {
        MergePart[] parts =
        [
            Part("表紙", 1),
            Part("議事録", 2, null, "第1回"),
            Part("資料", 1, null, "第1回"),
            Part("議事録", 3, null, "第2回"),
            Part("図", 1, null, "第2回", "添付"),
            Part("付録", 1),
        ];
        var bookmarks = PdfMerger.Bookmarks(parts, MergeOptions.None);
        Assert.Equal("表紙@0 第1回@1[議事録@1 資料@3] 第2回@4[議事録@4 添付@7[図@7]] 付録@8", Show(bookmarks));
        Assert.Equal(BookmarkKind.Folder, bookmarks[1].Kind);
        Assert.Equal(BookmarkKind.File, bookmarks[1].Children[0].Kind);
    }

    [Fact]
    public void Bookmarks_open_a_new_group_when_another_folder_comes_in_between()
    {
        MergePart[] parts = [Part("a", 1, null, "第1回"), Part("b", 1, null, "第2回"), Part("c", 1, null, "第1回")];
        Assert.Equal("第1回@0[a@0] 第2回@1[b@1] 第1回@2[c@2]", Show(PdfMerger.Bookmarks(parts, MergeOptions.None)));
    }

    [Fact]
    public void Bookmarks_tell_folders_apart_by_their_key_not_their_name()
    {
        // 名前を変えて同じ名前になっても、別のフォルダなら段を分ける
        MergePart[] parts =
        [
            Part("a", 1) with { Folders = [new("x/1", "資料")] },
            Part("b", 1) with { Folders = [new("x/2", "資料")] },
        ];
        Assert.Equal("資料@0[a@0] 資料@1[b@1]", Show(PdfMerger.Bookmarks(parts, MergeOptions.None)));
    }

    [Fact]
    public void Bookmarks_drop_a_folder_that_holds_only_a_file_of_the_same_name()
    {
        MergePart[] parts = [Part("見積", 1, null, "見積"), Part("議事録", 1, null, "第3回"), Part("図", 1, null, "第4回", "図")];
        var bookmarks = PdfMerger.Bookmarks(parts, MergeOptions.None);
        // 名前の違う 1 ファイルだけのフォルダは残す（「第3回」が消えると、ほかの回の議事録と見分けられない）
        Assert.Equal("見積@0 第3回@1[議事録@1] 第4回@2[図@2]", Show(bookmarks));
        Assert.Equal(0, bookmarks[0].Part);
    }

    [Fact]
    public void Bookmarks_keep_a_renamed_folder_even_when_its_new_name_matches_the_file()
    {
        // 名前を変えただけで段が消えると、木の上から元に戻せなくなる
        var part = Part("見積書", 1) with { Folders = [new BookmarkFolder("x/見積", "見積書")] };
        Assert.Equal("見積書@0[見積書@0]", Show(PdfMerger.Bookmarks([part], MergeOptions.None)));
    }

    [Theory]
    [InlineData(null, "報告書@0[報告書@0[背景@1]]")]
    [InlineData(1, "報告書@0")]
    public void Bookmarks_drop_a_lone_title_bookmark_left_after_cutting_the_depth(int? depth, string expected)
    {
        // 見出し 1 が表紙の題名だけの文書。1 段までにすると題名だけが残り、ファイルのしおりと重なる
        var part = Part("報告書", 3, [Node("報告書", 0, Node("背景", 1))]);
        Assert.Equal(expected, Show(PdfMerger.Bookmarks([part], MergeOptions.None with { OutlineDepth = depth })));
    }

    [Fact]
    public void Bookmarks_rename_and_drop_original_bookmarks_by_their_position()
    {
        // 「1.1」の名前を変え、「第2章」を外す
        var part = Part("資料", 4, Chapters) with
        {
            OutlineEdits = new Dictionary<string, string?> { ["0/0"] = "背景", ["1"] = null },
        };
        var bookmarks = PdfMerger.Bookmarks([part], MergeOptions.None);
        Assert.Equal("資料@0[第1章@0[背景@1[1.1.1@1]]]", Show(bookmarks));
        Assert.Equal("0/0", bookmarks[0].Children[0].Children[0].OutlinePath);
    }

    [Fact]
    public void Bookmarks_drop_the_children_of_a_dropped_bookmark_too()
    {
        var part = Part("資料", 4, Chapters) with { OutlineEdits = new Dictionary<string, string?> { ["0"] = null } };
        Assert.Equal("資料@0[第2章@3]", Show(PdfMerger.Bookmarks([part], MergeOptions.None)));
    }

    [Fact]
    public void Bookmarks_keep_edits_on_the_same_bookmark_when_pages_or_depth_change()
    {
        // 1 ページ目を外して「第1章」が付かなくなっても、「1.1」の手直しは「1.1」に当たる
        var part = new MergePart("unused.pdf", "資料", [1, 2, 3], Chapters)
        {
            OutlineEdits = new Dictionary<string, string?> { ["0/0"] = "背景" },
        };
        Assert.Equal("資料@0[背景@0[1.1.1@0] 第2章@2]", Show(PdfMerger.Bookmarks([part], MergeOptions.None)));
        Assert.Equal("資料@0[背景@0 第2章@2]", Show(PdfMerger.Bookmarks([part], MergeOptions.None with { OutlineDepth = 2 })));
    }

    [Fact]
    public void Bookmarks_place_added_bookmarks_among_the_originals_by_page()
    {
        // 第1章(0) 第2章(3) の間の 2 ページ目に足すと間に入り、同じページなら元のしおりの後ろに入る
        var part = Part("資料", 4, Chapters) with { Added = [new("図表", 2), new("補足", 3)] };
        var bookmarks = PdfMerger.Bookmarks([part], MergeOptions.None);
        Assert.Equal("資料@0[第1章@0[1.1@1[1.1.1@1]] 図表@2 第2章@3 補足@3]", Show(bookmarks));
        Assert.Equal(BookmarkKind.Added, bookmarks[0].Children[1].Kind);
        Assert.Equal(1, bookmarks[0].Children[3].AddedIndex);
    }

    [Theory]
    [InlineData(0, "資料@0[図表@2]")]
    [InlineData(1, "資料@0[第1章@0 図表@2 第2章@3]")]
    public void Bookmarks_keep_added_bookmarks_whatever_the_depth(int depth, string expected)
    {
        var part = Part("資料", 4, Chapters) with { Added = [new("図表", 2)] };
        Assert.Equal(expected, Show(PdfMerger.Bookmarks([part], MergeOptions.None with { OutlineDepth = depth })));
    }

    [Fact]
    public void Bookmarks_leave_out_added_bookmarks_on_pages_not_merged()
    {
        var part = new MergePart("unused.pdf", "スキャン", [0, 2], []) { Added = [new("表紙", 0), new("外したページ", 1), new("献立", 2)] };
        Assert.Equal("スキャン@0[表紙@0 献立@1]", Show(PdfMerger.Bookmarks([part], MergeOptions.None)));
    }

    [Fact]
    public void Merge_writes_added_bookmarks_to_the_pdf()
    {
        string a = _folder.Pdf("scan.pdf", 3);
        MergePart[] parts = [new(a, "スキャン", [0, 1, 2], []) { Added = [new("献立", 1), new("付録", 2)] }];
        string output = _folder.File("out.pdf");

        PdfMerger.Merge(parts, MergeOptions.None, output, ct: Ct);

        Assert.Equal("スキャン@0[献立@1 付録@2]", Show(PdfInfo.Read(output).Outline));
    }

    [Theory]
    // 資料@1[第1章@1[1.1@2[1.1.1@2]] 図表@3 第2章@4]、別のファイル b@5（目次が 1 ページ目）
    [InlineData(0, 1, "第1章")]   // ファイルと章が同じページなら、細かいほう
    [InlineData(0, 2, "1.1.1")]   // 同じページなら、いちばん細かい見出し
    [InlineData(0, 3, "図表")]    // 足したしおり
    [InlineData(0, 4, "第2章")]
    [InlineData(1, 5, "b")]
    [InlineData(1, 4, null)]      // そのファイルはまだ始まっていない
    public void Covering_finds_the_finest_bookmark_that_starts_at_or_before_the_page(int part, int page, string? expected)
    {
        MergePart[] parts = [Part("資料", 4, Chapters) with { Added = [new("図表", 2)] }, Part("b", 1)];
        var bookmarks = PdfMerger.Bookmarks(parts, MergeOptions.None with { TableOfContents = true });
        Assert.Equal(expected, DocsToPdf.Core.Bookmarks.Covering(bookmarks, part, page)?.Title);
    }

    [Fact]
    public void WriteContents_writes_only_the_contents_pages()
    {
        string a = _folder.Pdf("a.pdf", 2);
        MergePart[] parts = [new(a, "議事録", [0, 1], []), new(a, "資料", [0], [])];
        string output = _folder.File("contents.pdf");

        Assert.Equal(1, PdfMerger.WriteContents(parts, MergeOptions.None, output));
        Assert.Single(PdfInfo.Read(output).Sizes);
        Assert.Equal(0, PdfMerger.WriteContents([], MergeOptions.None, _folder.File("none.pdf")));
    }

    [Fact]
    public void Bookmarks_let_a_file_choose_its_own_depth()
    {
        MergePart[] parts =
        [
            Part("スライド", 4, Chapters) with { OutlineDepth = 0 },
            Part("報告", 4, Chapters),
            Part("仕様", 4, Chapters) with { OutlineDepth = MergePart.AllLevels },
        ];
        var bookmarks = PdfMerger.Bookmarks(parts, MergeOptions.None with { OutlineDepth = 1 });
        Assert.Equal("スライド@0 報告@4[第1章@4 第2章@7] 仕様@8[第1章@8[1.1@9[1.1.1@9]] 第2章@11]", Show(bookmarks));
        Assert.Equal([0, 1, 2], bookmarks.Select(b => b.Part));
    }

    [Fact]
    public void Bookmarks_put_the_table_of_contents_first_and_shift_the_files()
    {
        var bookmarks = PdfMerger.Bookmarks([Part("a", 2), Part("b", 1)], MergeOptions.None with { TableOfContents = true });
        Assert.Equal("目次@0 a@1 b@3", Show(bookmarks));
        Assert.Equal(BookmarkKind.Contents, bookmarks[0].Kind);
    }

    [Fact]
    public void Bookmarks_without_file_bookmarks_place_the_original_ones_at_the_top()
    {
        var bookmarks = PdfMerger.Bookmarks([Part("資料", 4, Chapters)], MergeOptions.None with { FileBookmarks = false, OutlineDepth = 1 });
        Assert.Equal("第1章@0 第2章@3", Show(bookmarks));
    }

    [Fact]
    public void Merge_writes_the_folder_groups_and_the_chosen_depth()
    {
        string a = _folder.Pdf("a.pdf", 2, ("一", 0), ("二", 1));
        string b = _folder.Pdf("b.pdf", 1);
        var (_, outlineA) = PdfInfo.Read(a);
        MergePart[] parts =
        [
            new(a, "議事録", [0, 1], outlineA) { Folders = In("第1回") },
            new(b, "資料", [0], []) { Folders = In("第1回") },
        ];
        string output = _folder.File("out.pdf");

        PdfMerger.Merge(parts, MergeOptions.None, output, ct: Ct);
        Assert.Equal("第1回@0[議事録@0[一@0 二@1] 資料@2]", Show(PdfInfo.Read(output).Outline));

        PdfMerger.Merge(parts, MergeOptions.None with { OutlineDepth = 0 }, output, ct: Ct);
        Assert.Equal("第1回@0[議事録@0 資料@2]", Show(PdfInfo.Read(output).Outline));
    }

    [Fact]
    public void Merge_writes_the_same_bookmarks_it_plans()
    {
        string a = _folder.Pdf("a.pdf", 4, ("一", 0), ("三", 2));
        string b = _folder.Pdf("b.pdf", 2);
        var (_, outlineA) = PdfInfo.Read(a);
        MergePart[] parts = [new(a, "a", [0, 2, 3], outlineA) { Folders = In("x") }, new(b, "b", [0, 1], []) { OutlineDepth = 0 }];
        var options = MergeOptions.None with { TableOfContents = true, DuplexBlanks = true };
        string output = _folder.File("out.pdf");

        PdfMerger.Merge(parts, options, output, ct: Ct);

        Assert.Equal(Show(PdfMerger.Bookmarks(parts, options)), Show(PdfInfo.Read(output).Outline));
    }

    [Fact]
    public void BookmarkFolders_leave_out_the_name_of_a_single_folder()
    {
        string root = Path.Combine(_folder.Path, "会議");
        string[] files = [Path.Combine(root, "表紙.pdf"), Path.Combine(root, "第1回", "a.pdf"), Path.Combine(root, "第1回", "添付", "b.pdf"), _folder.File("別.pdf")];
        var folders = BookmarkFolders.Of(files, [root, root, root, null]);
        Assert.Equal(["", "第1回", "第1回/添付", ""], folders.Select(f => string.Join("/", f.Select(x => x.Name))));
        Assert.Equal(Path.Combine(root, "第1回", "添付"), folders[2][1].Key);
    }

    [Fact]
    public void BookmarkFolders_treat_a_folder_named_with_two_dots_as_inside()
    {
        string root = Path.Combine(_folder.Path, "会議");
        var folders = BookmarkFolders.Of([Path.Combine(root, "..old", "a.pdf")], [root]);
        Assert.Equal("..old", folders[0].Single().Name);
    }

    [Fact]
    public void BookmarkFolders_count_every_added_folder_not_only_the_converted_files()
    {
        // 「第2回」のファイルがまだ変換中でも、2 つ入れたことに変わりはないので「第1回」の名前を段にする
        string first = Path.Combine(_folder.Path, "第1回"), second = Path.Combine(_folder.Path, "第2回");
        var folders = BookmarkFolders.Of([Path.Combine(first, "a.pdf")], [first], [first, second]);
        Assert.Equal("第1回", folders[0].Single().Name);
    }

    [Fact]
    public void BookmarkFolders_name_each_folder_when_several_were_added()
    {
        string first = Path.Combine(_folder.Path, "第1回"), second = Path.Combine(_folder.Path, "第2回");
        string[] files = [Path.Combine(first, "a.pdf"), Path.Combine(second, "資料", "b.pdf")];
        var folders = BookmarkFolders.Of(files, [first, second + Path.DirectorySeparatorChar]);
        Assert.Equal(["第1回", "第2回/資料"], folders.Select(f => string.Join("/", f.Select(x => x.Name))));
        Assert.Equal([first, second], folders.Select(f => f[0].Key));
    }

    [Fact]
    public void InputFiles_remember_the_folder_each_file_was_found_in()
    {
        string folder = Path.Combine(_folder.Path, "資料");
        Directory.CreateDirectory(Path.Combine(folder, "添付"));
        string inner = _folder.Pdf(Path.Combine("資料", "添付", "b.pdf"), 1);
        string direct = _folder.Pdf("直接.pdf", 1);

        var result = InputFiles.Expand([direct, folder], sortAll: false, ct: Ct);

        Assert.Equal([direct, inner], result.Files);
        Assert.Null(result.RootOf(direct));
        Assert.Equal(folder, result.RootOf(inner));
    }
}
