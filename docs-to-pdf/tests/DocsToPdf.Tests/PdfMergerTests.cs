using DocsToPdf.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Tests;

public sealed class PdfMergerTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    private static CancellationToken Ct => TestContext.Current.CancellationToken;

    private MergePart Part(string title, string path, IReadOnlyList<int>? pages = null)
    {
        var (sizes, outline) = PdfInfo.Read(path);
        return new MergePart(path, title, pages ?? Enumerable.Range(0, sizes.Count).ToArray(), outline);
    }

    private static string[] Titles(IEnumerable<OutlineNode> nodes) => nodes.Select(n => n.Title).ToArray();

    [Fact]
    public void Plan_places_files_one_after_another()
    {
        var (total, starts, blanks) = PdfMerger.Plan([2, 3, 1], MergeOptions.None);
        Assert.Equal(new[] { 0, 2, 5 }, starts);
        Assert.DoesNotContain(true, blanks);
        Assert.Equal(6, total);
    }

    [Fact]
    public void Plan_puts_the_table_of_contents_first()
    {
        var (total, starts, _) = PdfMerger.Plan([2, 1], MergeOptions.None with { TableOfContents = true });
        Assert.Equal(new[] { 1, 3 }, starts);
        Assert.Equal(4, total);
    }

    [Fact]
    public void Plan_adds_blanks_so_each_file_starts_on_an_odd_page()
    {
        var (total, starts, blanks) = PdfMerger.Plan([3, 2, 1], MergeOptions.None with { DuplexBlanks = true });
        Assert.Equal(new[] { 0, 4, 6 }, starts);
        Assert.Equal(new[] { false, true, false }, blanks);
        Assert.Equal(7, total);
    }

    [Fact]
    public void TocPageCount_grows_with_the_number_of_files()
    {
        Assert.Equal(1, PdfMerger.TocPageCount(1));
        int perPage = Enumerable.Range(1, 200).First(n => PdfMerger.TocPageCount(n + 1) == 2);
        Assert.Equal(2, PdfMerger.TocPageCount(perPage * 2));
        Assert.Equal(3, PdfMerger.TocPageCount(perPage * 2 + 1));
    }

    [Fact]
    public void Merge_joins_pages_and_adds_a_bookmark_for_each_file()
    {
        var a = Part("報告書", _folder.Pdf("a.pdf", 2));
        var b = Part("売上", _folder.Pdf("b.pdf", 3));
        string output = _folder.File("out.pdf");

        int pages = PdfMerger.Merge([a, b], MergeOptions.None, output, ct: Ct);

        var (sizes, outline) = PdfInfo.Read(output);
        Assert.Equal(5, pages);
        Assert.Equal(5, sizes.Count);
        Assert.Equal(new[] { "報告書", "売上" }, Titles(outline));
        Assert.Equal(new[] { 0, 2 }, outline.Select(n => n.PageIndex));
    }

    [Fact]
    public void Merge_nests_the_original_bookmarks_under_the_file()
    {
        var part = Part("資料", _folder.Pdf("a.pdf", 4, ("第一", 0), ("第二", 2)));
        string output = _folder.File("out.pdf");

        PdfMerger.Merge([Part("表紙", _folder.Pdf("c.pdf", 1)), part], MergeOptions.None, output, ct: Ct);

        var file = PdfInfo.Read(output).Outline[1];
        Assert.Equal("資料", file.Title);
        Assert.Equal(new[] { "第一", "第二" }, Titles(file.Children));
        Assert.Equal(new[] { 1, 3 }, file.Children.Select(n => n.PageIndex));
    }

    [Fact]
    public void Merge_skips_a_single_bookmark_that_only_points_at_the_first_page()
    {
        var part = Part("資料", _folder.Pdf("a.pdf", 2, ("資料の題", 0)));
        string output = _folder.File("out.pdf");

        PdfMerger.Merge([part], MergeOptions.None, output, ct: Ct);

        Assert.Empty(PdfInfo.Read(output).Outline.Single().Children);
    }

    [Fact]
    public void Merge_uses_only_the_chosen_pages_and_drops_bookmarks_to_excluded_pages()
    {
        string path = _folder.Pdf("a.pdf", 5, ("一", 0), ("三", 2), ("五", 4));
        var part = Part("資料", path, [0, 3, 4]);
        string output = _folder.File("out.pdf");

        int pages = PdfMerger.Merge([part], MergeOptions.None, output, ct: Ct);

        var file = PdfInfo.Read(output).Outline.Single();
        Assert.Equal(3, pages);
        Assert.Equal(new[] { "一", "五" }, Titles(file.Children));
        Assert.Equal(new[] { 0, 2 }, file.Children.Select(n => n.PageIndex));
    }

    [Fact]
    public void Merge_with_all_options_produces_the_planned_page_count()
    {
        var parts = new[] { Part("一", _folder.Pdf("a.pdf", 3)), Part("二", _folder.Pdf("b.pdf", 2)) };
        var options = new MergeOptions(PageNumbers: true, TableOfContents: true, DuplexBlanks: true, "まとめ");
        string output = _folder.File("out.pdf");

        int pages = PdfMerger.Merge(parts, options, output, ct: Ct);

        var (total, _, _) = PdfMerger.Plan(parts.Select(p => p.Pages.Count).ToList(), options);
        Assert.Equal(total, pages);
        var outline = PdfInfo.Read(output).Outline;
        Assert.Equal(new[] { "目次", "一", "二" }, Titles(outline));
        using var doc = PdfReader.Open(output, PdfDocumentOpenMode.Import);
        Assert.Equal("まとめ", doc.Info.Title);
        Assert.True(doc.Pages[0].Annotations.Count >= 2, "目次の行にリンクがない");
    }

    [Fact]
    public void Merge_stamps_page_numbers_on_rotated_pages()
    {
        string path = _folder.File("rotated.pdf");
        using (var doc = new PdfDocument())
        {
            foreach (int rotate in new[] { 0, 90, 180, 270 }) doc.AddPage().Rotate = rotate;
            doc.Save(path);
        }
        string output = _folder.File("out.pdf");

        int pages = PdfMerger.Merge([Part("回転", path)], MergeOptions.None with { PageNumbers = true }, output, ct: Ct);

        Assert.Equal(4, pages);
    }

    [Fact]
    public void Merge_turns_the_pages_the_user_rotated()
    {
        string path = _folder.File("mixed.pdf");
        using (var doc = new PdfDocument())
        {
            doc.AddPage();                 // 縦
            doc.AddPage().Rotate = 90;     // 元から回っている
            doc.AddPage();
            doc.Save(path);
        }
        var rotations = new Dictionary<int, int> { [0] = 90, [1] = 270 };
        var part = Part("回転", path) with { Rotations = rotations };
        string output = _folder.File("out.pdf");

        PdfMerger.Merge([part], MergeOptions.None with { PageNumbers = true }, output, ct: Ct);

        using var result = PdfReader.Open(output, PdfDocumentOpenMode.Import);
        Assert.Equal(new[] { 90, 0, 0 }, result.Pages.Cast<PdfPage>().Select(PdfInfo.Rotation));
        var sizes = PdfInfo.Read(output).Sizes;
        Assert.True(sizes[0].Width > sizes[0].Height, "回した 1 ページ目が横長になっていない");
        Assert.True(sizes[2].Width < sizes[2].Height, "回していない 3 ページ目が縦長のままでない");
    }

    [Fact]
    public void Merge_turns_the_duplex_blank_to_match_a_rotated_first_page()
    {
        var first = Part("一", _folder.Pdf("a.pdf", 1));
        var second = Part("二", _folder.Pdf("b.pdf", 1)) with { Rotations = new Dictionary<int, int> { [0] = 90 } };
        string output = _folder.File("out.pdf");

        PdfMerger.Merge([first, second], MergeOptions.None with { DuplexBlanks = true }, output, ct: Ct);

        var sizes = PdfInfo.Read(output).Sizes;
        Assert.Equal(3, sizes.Count);
        Assert.True(sizes[1].Width > sizes[1].Height, "白紙が、回したページと同じ横向きになっていない");
    }

    [Fact]
    public void Merge_leaves_the_existing_file_untouched_when_cancelled()
    {
        string output = _folder.File("out.pdf");
        File.WriteAllText(output, "前の内容");
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            PdfMerger.Merge([Part("一", _folder.Pdf("a.pdf", 1))], MergeOptions.None, output, ct: cancel.Token));

        Assert.Equal("前の内容", File.ReadAllText(output));
        Assert.Empty(Directory.GetFiles(_folder.Path, "*.tmp"));
    }
}
