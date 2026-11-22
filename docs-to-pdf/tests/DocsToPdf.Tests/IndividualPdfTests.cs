using DocsToPdf.Core;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Tests;

public class IndividualPdfTests
{
    private static readonly string Docs = Path.Combine(Path.GetTempPath(), "docs");
    private static string In(string name) => Path.Combine(Docs, name);

    [Fact]
    public void Names_each_pdf_after_its_file_next_to_it()
    {
        var names = IndividualPdf.FileNames([In("見積書.docx"), In("請求書.xlsx")], Path.GetDirectoryName!);

        Assert.Equal([In("見積書.pdf"), In("請求書.pdf")], names);
    }

    [Fact]
    public void Adds_the_extension_when_two_files_would_get_the_same_name()
    {
        var names = IndividualPdf.FileNames([In("報告.docx"), In("報告.XLSX"), In("表紙.pptx")], Path.GetDirectoryName!);

        Assert.Equal([In("報告 (docx).pdf"), In("報告 (xlsx).pdf"), In("表紙.pdf")], names);
    }

    [Fact]
    public void Never_writes_over_an_input_pdf()
    {
        var names = IndividualPdf.FileNames([In("図面.pdf"), In("図面 (pdf).pdf")], Path.GetDirectoryName!);

        Assert.DoesNotContain(In("図面.pdf"), names);
        Assert.DoesNotContain(In("図面 (pdf).pdf"), names);
        Assert.Equal(2, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void Avoids_reserved_paths_such_as_files_that_could_not_be_converted()
    {
        // 一覧に残っている（PDF にできなかった）報告.pdf には書かない
        var names = IndividualPdf.FileNames([In("報告.docx")], Path.GetDirectoryName!, reserved: [In("報告.pdf")]);

        Assert.Equal([In("報告 (docx).pdf")], names);
    }

    [Fact]
    public void Same_names_in_different_folders_do_not_clash_when_written_next_to_each_file()
    {
        string a = Path.Combine(Docs, "a", "報告.docx"), b = Path.Combine(Docs, "b", "報告.docx");

        var names = IndividualPdf.FileNames([a, b], Path.GetDirectoryName!);

        Assert.Equal([Path.Combine(Docs, "a", "報告.pdf"), Path.Combine(Docs, "b", "報告.pdf")], names);
    }

    [Fact]
    public void Same_names_from_different_folders_get_numbers_in_one_output_folder()
    {
        string a = Path.Combine(Docs, "a", "報告.docx"), b = Path.Combine(Docs, "b", "報告.docx");

        var names = IndividualPdf.FileNames([a, b], _ => Docs);

        Assert.Equal([In("報告 (docx).pdf"), In("報告 (docx) (2).pdf")], names);
    }

    [Fact]
    public void Write_keeps_only_the_chosen_pages_with_their_rotation_and_bookmarks()
    {
        using var folder = new TestFolder();
        string source = folder.Pdf("source.pdf", 4, ("一", 0), ("三", 2));
        string output = folder.File("out.pdf");
        var part = new MergePart(source, "資料", [0, 2], PdfInfo.Read(source).Outline, new Dictionary<int, int> { [2] = 90 });

        int pages = IndividualPdf.Write(part, output, pageNumbers: false, TestContext.Current.CancellationToken);

        Assert.Equal(2, pages);
        var (sizes, outline) = PdfInfo.Read(output);
        Assert.Equal(2, sizes.Count);
        // 1 ファイルだけなので「資料」のしおりで包まず、元のしおりを一番上に置く
        Assert.Equal(["一", "三"], outline.Select(o => o.Title));
        using var doc = PdfReader.Open(output, PdfDocumentOpenMode.Import);
        Assert.Equal(90, doc.Pages[1].Rotate);
        Assert.Equal("資料", doc.Info.Title);
    }
}
