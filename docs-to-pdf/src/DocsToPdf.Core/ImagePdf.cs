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
    /// <summary>小さく作るときの解像度（ページに載せた大きさでの dpi）。画面で読むには足り、A4 の写真で数百 KB に収まる。</summary>
    public const double SmallDpi = 150;

    /// <param name="small">載せる大きさで SmallDpi を超える画像は縮め、JPEG は画質を下げて書き直す。</param>
    public static void Create(string source, string pdfPath, bool small = false)
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
                bool passThrough = !small && frames == 1 && orientation == RotateFlipType.RotateNoneFlipNone
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
                        var stream = new MemoryStream();
                        streams.Add(stream);
                        var upright = new Bitmap(image);
                        try
                        {
                            upright.RotateFlip(orientation);
                            bool turned = upright.Width != image.Width;
                            naturalWidth = upright.Width * 72.0 / (turned ? dpiY : dpiX);
                            naturalHeight = upright.Height * 72.0 / (turned ? dpiX : dpiY);
                            if (small && Shrunk(upright, naturalWidth, naturalHeight) is { } shrunk)
                            {
                                upright.Dispose();
                                upright = shrunk;
                            }

                            if (image.RawFormat.Equals(ImageFormat.Jpeg))
                                SaveJpeg(upright, stream, small ? 75L : 92L);
                            else
                                upright.Save(stream, ImageFormat.Png);
                        }
                        finally
                        {
                            upright.Dispose();
                        }
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

    /// <summary>ページに載せたときの大きさ（ポイント）。大きければページに収め、小さければ実寸のまま。</summary>
    private static (double PageWidth, double PageHeight, double Width, double Height) Layout(double naturalWidth, double naturalHeight)
    {
        bool landscape = naturalWidth > naturalHeight;
        double pageWidth = landscape ? A4Long : A4Short;
        double pageHeight = landscape ? A4Short : A4Long;
        double scale = Math.Min(1.0, Math.Min(pageWidth / naturalWidth, pageHeight / naturalHeight));
        return (pageWidth, pageHeight, naturalWidth * scale, naturalHeight * scale);
    }

    /// <summary>載せる大きさで SmallDpi を超える画像を、SmallDpi まで縮めたもの。縮める必要がなければ null。</summary>
    private static Bitmap? Shrunk(Bitmap bitmap, double naturalWidth, double naturalHeight)
    {
        var (_, _, width, _) = Layout(naturalWidth, naturalHeight);
        double ratio = width / 72.0 * SmallDpi / bitmap.Width;
        if (ratio >= 0.9) return null; // ほとんど変わらないなら、縮めて画質を落とすだけになる
        int w = Math.Max(1, (int)Math.Round(bitmap.Width * ratio));
        int h = Math.Max(1, (int)Math.Round(bitmap.Height * ratio));
        // ページに載せる大きさは、縮める前の実寸から決める（AddPage に渡す naturalWidth / naturalHeight）
        var result = new Bitmap(w, h);
        using var g = Graphics.FromImage(result);
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
        g.DrawImage(bitmap, 0, 0, w, h);
        return result;
    }

    private static void AddPage(PdfDocument doc, XImage image, double naturalWidth, double naturalHeight)
    {
        var (pageWidth, pageHeight, width, height) = Layout(naturalWidth, naturalHeight);

        var page = doc.AddPage();
        page.Width = XUnit.FromPoint(pageWidth);
        page.Height = XUnit.FromPoint(pageHeight);

        using var g = XGraphics.FromPdfPage(page);
        g.DrawImage(image, (pageWidth - width) / 2, (pageHeight - height) / 2, width, height);
    }

    private static void SaveJpeg(Bitmap bitmap, Stream stream, long quality)
    {
        var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
        using var parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);
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
