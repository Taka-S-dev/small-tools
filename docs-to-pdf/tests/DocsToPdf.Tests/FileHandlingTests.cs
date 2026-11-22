using System.Drawing;
using System.Drawing.Imaging;
using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class FileHandlingTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    [Theory]
    [InlineData("a.docx", "Word")]
    [InlineData("a.RTF", "Word")]
    [InlineData("a.xlsb", "Excel")]
    [InlineData("a.ppsx", "PowerPoint")]
    [InlineData("a.pdf", "Pdf")]
    [InlineData("a.tiff", "Image")]
    public void TryGetKind_recognises_supported_files(string name, string expected)
    {
        // FileKind は internal なので、public なテストの引数には名前で渡す
        Assert.True(FileEntry.TryGetKind(name, out var kind));
        Assert.Equal(Enum.Parse<FileKind>(expected), kind);
    }

    [Theory]
    [InlineData("a.txt")]
    [InlineData("a")]
    [InlineData("~$報告書.docx")]
    public void TryGetKind_rejects_other_files_and_office_lock_files(string name) =>
        Assert.False(FileEntry.TryGetKind(name, out _));

    [Fact]
    public void NaturalStringComparer_compares_numbers_by_value()
    {
        var names = new List<string> { "資料10.pdf", "資料2.pdf", "資料1.pdf" };
        names.Sort(NaturalStringComparer.Instance);
        Assert.Equal(new[] { "資料1.pdf", "資料2.pdf", "資料10.pdf" }, names);
    }

    [Theory]
    [InlineData(400, 300, true)]
    [InlineData(300, 400, false)]
    public void ImagePdf_puts_one_image_on_an_A4_page_turned_to_match(int width, int height, bool landscape)
    {
        string image = _folder.File("image.png");
        using (var bitmap = new Bitmap(width, height))
            bitmap.Save(image, ImageFormat.Png);
        string pdf = _folder.File("image.pdf");

        ImagePdf.Create(image, pdf);

        var page = PdfInfo.Read(pdf).Sizes.Single();
        Assert.Equal(landscape, page.Width > page.Height);
        Assert.Equal(841.89, Math.Max(page.Width, page.Height), 1);
    }

    [Fact]
    public void ImagePdf_writes_every_page_of_a_multi_page_tiff()
    {
        string tiff = _folder.File("pages.tif");
        var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Tiff.Guid);
        using (var first = new Bitmap(100, 140))
        using (var second = new Bitmap(100, 140))
        using (var parameters = new EncoderParameters(1))
        {
            parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
            first.Save(tiff, codec, parameters);
            parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
            first.SaveAdd(second, parameters);
            parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
            first.SaveAdd(parameters);
        }
        string pdf = _folder.File("pages.pdf");

        ImagePdf.Create(tiff, pdf);

        Assert.Equal(2, PdfInfo.Read(pdf).Sizes.Count);
    }
}
