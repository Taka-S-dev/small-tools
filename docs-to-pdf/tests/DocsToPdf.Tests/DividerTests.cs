using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class DividerTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    private static CancellationToken Ct => TestContext.Current.CancellationToken;

    private static readonly MergeOptions WithDividers = MergeOptions.None with { Dividers = true };

    private static MergePart Part(string title, int pages, params string[] folders) =>
        new("unused.pdf", title, Enumerable.Range(0, pages).ToArray(), [])
        {
            Folders = folders.Select((name, d) => new BookmarkFolder(string.Join("/", folders[..(d + 1)]), name)).ToArray(),
        };

    [Fact]
    public void DividerTitles_puts_one_before_each_top_folder()
    {
        MergePart[] parts = [Part("a", 1, "契約"), Part("b", 1, "契約", "旧"), Part("c", 1, "見積"), Part("d", 1)];
        Assert.Equal(["契約", null, "見積", "d"], PdfMerger.DividerTitles(parts, WithDividers));
    }

    [Fact]
    public void DividerTitles_puts_one_before_each_file_when_not_grouped_by_folder()
    {
        MergePart[] parts = [Part("a", 1, "契約"), Part("b", 1, "契約")];
        Assert.Equal(["a", "b"], PdfMerger.DividerTitles(parts, WithDividers with { FileBookmarks = false }));
        Assert.Equal([null, null], PdfMerger.DividerTitles(parts, MergeOptions.None));
    }

    [Fact]
    public void DividerTitles_skips_files_without_pages()
    {
        MergePart[] parts = [Part("a", 0), Part("b", 2)];
        Assert.Equal([null, "b"], PdfMerger.DividerTitles(parts, WithDividers));
    }

    [Fact]
    public void Plan_gives_each_divider_a_page_before_its_file()
    {
        var plan = PdfMerger.Plan([2, 3], WithDividers with { TableOfContents = true }, [true, true]);
        Assert.Equal([1, 4], plan.Dividers);
        Assert.Equal([2, 5], plan.Starts);
        Assert.Equal(8, plan.TotalPages);
        Assert.Equal(2, plan.DividerCount);
    }

    [Fact]
    public void Plan_puts_dividers_on_the_front_and_leaves_their_back_blank_for_duplex()
    {
        // 3 ページのあとの中表紙は、白紙を挟んで表（偶数の位置）に置き、その裏も白紙にする
        var plan = PdfMerger.Plan([3, 1], WithDividers with { DuplexBlanks = true }, [true, true]);
        Assert.Equal([0, 6], plan.Dividers);
        Assert.Equal([2, 8], plan.Starts);
        Assert.Equal([true, true], plan.BlankBefore);
        Assert.Equal(3, plan.BlankPages);
        Assert.Equal(9, plan.TotalPages);
    }

    [Fact]
    public void Merge_adds_divider_pages_and_points_top_folders_at_them()
    {
        var a = _folder.Pdf("a.pdf", 2);
        var b = _folder.Pdf("b.pdf", 1);
        MergePart[] parts =
        [
            new(a, "A", [0, 1], []) { Folders = [new BookmarkFolder("x/契約", "契約")] },
            new(b, "B", [0], []),
        ];
        var options = WithDividers with { PageNumbers = true };
        string output = _folder.File("out.pdf");

        int pages = PdfMerger.Merge(parts, options, output, ct: Ct);

        Assert.Equal(5, pages); // 中表紙・A 2 ページ・中表紙・B
        Assert.Equal(PdfMerger.Plan(parts, options).TotalPages, pages);
        var outline = PdfInfo.Read(output).Outline;
        Assert.Equal(["契約", "B"], outline.Select(n => n.Title));
        Assert.Equal(0, outline[0].PageIndex);              // フォルダの段は中表紙を指す
        Assert.Equal(1, outline[0].Children[0].PageIndex);  // ファイルのしおりは中身の最初のページ
        Assert.Equal(4, outline[1].PageIndex);
    }

    [Fact]
    public void Bookmarks_preview_matches_the_merged_pages()
    {
        MergePart[] parts = [Part("A", 2, "契約"), Part("B", 1)];
        var tree = PdfMerger.Bookmarks(parts, WithDividers);
        Assert.Equal(0, tree[0].Page);
        Assert.Equal(1, tree[0].Children[0].Page);
        Assert.Equal(4, tree[1].Page);
    }
}
