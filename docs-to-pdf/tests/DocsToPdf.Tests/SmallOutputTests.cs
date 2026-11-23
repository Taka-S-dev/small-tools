using System.Drawing;
using System.Drawing.Imaging;
using DocsToPdf.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Tests;

public sealed class SmallOutputTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    /// <summary>細かい模様の大きな写真（JPEG）。縮めると目に見えてファイルが小さくなる。</summary>
    private string Photo(int width, int height)
    {
        string path = _folder.File($"photo{width}.jpg");
        using var bitmap = new Bitmap(width, height);
        var random = new Random(1);
        for (int y = 0; y < height; y += 2)
            for (int x = 0; x < width; x += 2)
                bitmap.SetPixel(x, y, Color.FromArgb(random.Next(256), (x * 255) / width, (y * 255) / height));
        bitmap.Save(path, ImageFormat.Jpeg);
        return path;
    }

    private static (int Width, int Height) FirstImage(string pdf)
    {
        using var doc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
        var objects = doc.Pages[0].Elements.GetDictionary("/Resources")?.Elements.GetDictionary("/XObject") ?? throw new InvalidOperationException("画像がない");
        var image = (objects.Elements.Values.First() as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary
            ?? throw new InvalidOperationException("画像が読めない");
        return (image.Elements.GetInteger("/Width"), image.Elements.GetInteger("/Height"));
    }

    [Fact]
    public void ImagePdf_small_shrinks_a_large_photo_but_keeps_the_page()
    {
        string photo = Photo(3000, 2000);
        string normal = _folder.File("normal.pdf"), small = _folder.File("small.pdf");

        ImagePdf.Create(photo, normal);
        ImagePdf.Create(photo, small, small: true);

        Assert.True(new FileInfo(small).Length < new FileInfo(normal).Length / 2);
        Assert.Equal(PdfInfo.Read(normal).Sizes, PdfInfo.Read(small).Sizes);
        // A4 の横向きいっぱい（約 11.7 インチ）に 150dpi なら、横はおよそ 1754 ピクセル
        Assert.InRange(FirstImage(small).Width, 1700, 1800);
    }

    [Fact]
    public void PdfImages_Shrink_rewrites_large_jpegs_and_keeps_the_pages()
    {
        string pdf = _folder.File("photo.pdf");
        ImagePdf.Create(Photo(3000, 2000), pdf); // 元の JPEG をそのまま埋め込む
        long before = new FileInfo(pdf).Length;
        var sizes = PdfInfo.Read(pdf).Sizes;

        Assert.Equal(1, PdfImages.Shrink(pdf, 150));

        Assert.True(new FileInfo(pdf).Length < before / 2);
        Assert.Equal(sizes, PdfInfo.Read(pdf).Sizes);
        var (width, height) = FirstImage(pdf);
        Assert.InRange(width, 1700, 1800);
        Assert.Equal(2.0 / 3, (double)height / width, 2);
    }

    [Fact]
    public void PdfImages_Shrink_leaves_small_images_alone()
    {
        string pdf = _folder.File("photo.pdf");
        ImagePdf.Create(Photo(600, 400), pdf);
        byte[] before = File.ReadAllBytes(pdf);

        Assert.Equal(0, PdfImages.Shrink(pdf, 150));
        Assert.Equal(before, File.ReadAllBytes(pdf));
    }
}
