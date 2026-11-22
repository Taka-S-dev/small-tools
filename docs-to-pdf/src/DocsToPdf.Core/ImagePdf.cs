using System.Drawing;
using System.Drawing.Imaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace DocsToPdf.Core;

/// <summary>
/// 画像を PDF にする。1 枚を A4 の 1 ページに載せ、横長の画像は横向きのページにする。
/// 大きい画像はページいっぱいに縮め、小さい画像は拡大せず実寸のまま中央に置く。
/// 複数ページの TIFF は全ページを出し、スマホ写真の向き（EXIF）も反映する。
/// </summary>
internal static class ImagePdf
{
    private const double A4Short = 595.276, A4Long = 841.89;
    private const int ExifOrientation = 0x0112;

    public static void Create(string source, string pdfPath)
    {
        using var doc = new PdfDocument();
        var streams = new List<Stream>();
        try
        {
            using (var image = Image.FromFile(source))
            {
                var orientation = ReadOrientation(image);
                bool multiPage = image.FrameDimensionsList.Contains(FrameDimension.Page.Guid);
                int frames = multiPage ? image.GetFrameCount(FrameDimension.Page) : 1;
                bool passThrough = frames == 1 && orientation == RotateFlipType.RotateNoneFlipNone
                    && (image.RawFormat.Equals(ImageFormat.Jpeg) || image.RawFormat.Equals(ImageFormat.Png));

                for (int frame = 0; frame < frames; frame++)
                {
                    if (multiPage) image.SelectActiveFrame(FrameDimension.Page, frame);
                    double dpiX = image.HorizontalResolution > 1 ? image.HorizontalResolution : 96;
                    double dpiY = image.VerticalResolution > 1 ? image.VerticalResolution : 96;

                    XImage xImage;
                    double naturalWidth, naturalHeight;
                    if (passThrough)
                    {
                        // 元の JPEG / PNG をそのまま埋め込む（再圧縮で画質を落とさない）
                        naturalWidth = image.Width * 72.0 / dpiX;
                        naturalHeight = image.Height * 72.0 / dpiY;
                        xImage = XImage.FromFile(source);
                    }
                    else
                    {
                        using var upright = new Bitmap(image);
                        upright.RotateFlip(orientation);
                        bool turned = upright.Width != image.Width;
                        naturalWidth = upright.Width * 72.0 / (turned ? dpiY : dpiX);
                        naturalHeight = upright.Height * 72.0 / (turned ? dpiX : dpiY);

                        var stream = new MemoryStream();
                        streams.Add(stream);
                        if (image.RawFormat.Equals(ImageFormat.Jpeg))
                            SaveJpeg(upright, stream);
                        else
                            upright.Save(stream, ImageFormat.Png);
                        stream.Position = 0;
                        xImage = XImage.FromStream(stream);
                    }

                    using (xImage)
                        AddPage(doc, xImage, naturalWidth, naturalHeight);
                }
            }
            doc.Save(pdfPath);
        }
        finally
        {
            foreach (var stream in streams) stream.Dispose();
        }
    }

    private static void AddPage(PdfDocument doc, XImage image, double naturalWidth, double naturalHeight)
    {
        bool landscape = naturalWidth > naturalHeight;
        double pageWidth = landscape ? A4Long : A4Short;
        double pageHeight = landscape ? A4Short : A4Long;

        var page = doc.AddPage();
        page.Width = XUnit.FromPoint(pageWidth);
        page.Height = XUnit.FromPoint(pageHeight);

        double scale = Math.Min(1.0, Math.Min(pageWidth / naturalWidth, pageHeight / naturalHeight));
        double width = naturalWidth * scale, height = naturalHeight * scale;
        using var g = XGraphics.FromPdfPage(page);
        g.DrawImage(image, (pageWidth - width) / 2, (pageHeight - height) / 2, width, height);
    }

    private static void SaveJpeg(Bitmap bitmap, Stream stream)
    {
        var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
        using var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(Encoder.Quality, 92L);
        bitmap.Save(stream, codec, parameters);
    }

    private static RotateFlipType ReadOrientation(Image image)
    {
        try
        {
            if (!image.PropertyIdList.Contains(ExifOrientation)) return RotateFlipType.RotateNoneFlipNone;
            var value = image.GetPropertyItem(ExifOrientation)?.Value;
            if (value is null || value.Length < 2) return RotateFlipType.RotateNoneFlipNone;
            return BitConverter.ToUInt16(value, 0) switch
            {
                2 => RotateFlipType.RotateNoneFlipX,
                3 => RotateFlipType.Rotate180FlipNone,
                4 => RotateFlipType.Rotate180FlipX,
                5 => RotateFlipType.Rotate90FlipX,
                6 => RotateFlipType.Rotate90FlipNone,
                7 => RotateFlipType.Rotate270FlipX,
                8 => RotateFlipType.Rotate270FlipNone,
                _ => RotateFlipType.RotateNoneFlipNone,
            };
        }
        catch
        {
            return RotateFlipType.RotateNoneFlipNone;
        }
    }
}
