using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Core;

/// <summary>
/// PDF の中の写真（JPEG）を縮めて入れ直し、ファイルを小さくする。
/// Office の「最小サイズ」で書き出せない PowerPoint の PDF に使う。
/// 取り違えると絵が壊れるので、扱いが確かな形（RGB かグレーの 8 ビットの JPEG で、Decode の指定がないもの）だけを縮め、ほかは触らない。
/// </summary>
internal static class PdfImages
{
    /// <summary>
    /// どの画像も、その PDF の一番大きいページいっぱいに載せたときに dpi を超えない大きさまで縮める
    /// （どの大きさで載っているかまでは調べないので、ページより大きく載ることはない、という上限で決める）。
    /// </summary>
    /// <returns>縮めた画像の数。</returns>
    public static int Shrink(string pdfPath, double dpi, long quality = 75)
    {
        int shrunk = 0;
        using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
        {
            double longest = doc.Pages.Cast<PdfPage>().Max(p => Math.Max(p.Width.Point, p.Height.Point));
            int limit = (int)Math.Ceiling(longest / 72.0 * dpi);
            var seen = new HashSet<PdfDictionary>(ReferenceEqualityComparer.Instance);
            foreach (PdfPage page in doc.Pages)
                shrunk += ShrinkIn(page.Elements.GetDictionary("/Resources"), limit, quality, seen);
            if (shrunk == 0) return 0;
            doc.Save(pdfPath);
        }
        return shrunk;
    }

    private static int ShrinkIn(PdfDictionary? resources, int limit, long quality, HashSet<PdfDictionary> seen)
    {
        var objects = resources?.Elements.GetDictionary("/XObject");
        if (objects is null) return 0;
        int shrunk = 0;
        foreach (var key in objects.Elements.Keys.ToList())
        {
            if ((objects.Elements[key] is PdfReference r ? r.Value : objects.Elements[key]) is not PdfDictionary item || !seen.Add(item))
                continue;
            string subtype = item.Elements.GetName("/Subtype");
            if (subtype == "/Form")
                shrunk += ShrinkIn(item.Elements.GetDictionary("/Resources"), limit, quality, seen); // 図形のまとまりの中の画像
            else if (subtype == "/Image" && TryShrink(item, limit, quality))
                shrunk++;
        }
        return shrunk;
    }

    private static bool TryShrink(PdfDictionary image, int limit, long quality)
    {
        try
        {
            if (image.Elements.GetName("/Filter") != "/DCTDecode") return false; // 1 段の JPEG だけ
            if (image.Elements.ContainsKey("/Decode") || image.Elements.GetInteger("/BitsPerComponent") != 8) return false;
            if (image.Elements.GetBoolean("/ImageMask")) return false;
            bool gray = image.Elements.GetName("/ColorSpace") == "/DeviceGray";
            if (!gray && image.Elements.GetName("/ColorSpace") != "/DeviceRGB" && !IsIccRgb(image)) return false;

            int width = image.Elements.GetInteger("/Width"), height = image.Elements.GetInteger("/Height");
            double ratio = (double)limit / Math.Max(width, height);
            if (ratio >= 0.9 || image.Stream?.Value is not { Length: > 0 } bytes) return false; // ほとんど変わらないなら触らない

            using var source = new Bitmap(new MemoryStream(bytes));
            if (source.Width != width || source.Height != height) return false;
            int w = Math.Max(1, (int)Math.Round(width * ratio)), h = Math.Max(1, (int)Math.Round(height * ratio));
            using var resized = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(resized))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(source, 0, 0, w, h);
            }
            using var output = new MemoryStream();
            var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            using (var parameters = new EncoderParameters(1))
            {
                parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                resized.Save(output, codec, parameters);
            }
            if (output.Length >= bytes.Length) return false; // 小さくならないなら元のまま

            image.Stream.Value = output.ToArray();
            image.Elements.SetInteger("/Width", w);
            image.Elements.SetInteger("/Height", h);
            // 書き直した JPEG はいつも RGB
            if (gray) image.Elements.SetName("/ColorSpace", "/DeviceRGB");
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or ExternalException or InvalidOperationException)
        {
            return false; // 読めない JPEG は元のまま
        }
    }

    /// <summary>色の決まりが RGB の ICC プロファイル（[/ICCBased ストリーム] で /N が 3）か。</summary>
    private static bool IsIccRgb(PdfDictionary image)
    {
        var item = image.Elements["/ColorSpace"];
        if ((item is PdfReference r ? r.Value : item) is not PdfArray { Elements.Count: 2 } array) return false;
        if (array.Elements.GetName(0) != "/ICCBased") return false;
        var profile = array.Elements[1] is PdfReference p ? p.Value as PdfDictionary : array.Elements[1] as PdfDictionary;
        return profile?.Elements.GetInteger("/N") == 3;
    }
}
