using System.Buffers.Binary;
using PdfSharp.Fonts;

namespace DocsToPdf.Core;

/// <summary>
/// ページ番号と目次の文字を描くためのフォント。Windows に入っている日本語フォントを使う。
/// PDFsharp は TTC（複数のフォントを 1 つにまとめた形式）を読めないため、中から 1 書体を取り出して渡す。
/// </summary>
internal sealed class PdfFonts : IFontResolver
{
    public const string Family = "DocsToPdfText";

    private static readonly (string File, int Index)[] Candidates =
    [
        ("YuGothM.ttc", 0),
        ("meiryo.ttc", 0),
        ("msgothic.ttc", 0),
        ("yumin.ttf", 0),
        ("segoeui.ttf", 0),
        ("arial.ttf", 0),
    ];

    private static readonly object Gate = new();
    private static bool _registered;
    private byte[]? _font;

    /// <summary>PDFsharp にこのフォントを登録する。何度呼んでもよい。</summary>
    public static void EnsureRegistered()
    {
        lock (Gate)
        {
            if (_registered) return;
            GlobalFontSettings.FontResolver = new PdfFonts();
            _registered = true;
        }
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic) => new(Family);

    public byte[]? GetFont(string faceName)
    {
        lock (Gate)
        {
            return _font ??= Load();
        }
    }

    private static byte[] Load()
    {
        string fonts = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        foreach (var (file, index) in Candidates)
        {
            string path = Path.Combine(fonts, file);
            try
            {
                if (File.Exists(path))
                    return ExtractFace(File.ReadAllBytes(path), index);
            }
            catch
            {
                // 読めないフォントは飛ばして次を試す
            }
        }
        throw new InvalidOperationException("文字を描くためのフォントが見つかりません。");
    }

    /// <summary>TTC から index 番目の書体を取り出し、単独の TrueType フォントとして組み直す。TTC でなければそのまま返す。</summary>
    private static byte[] ExtractFace(byte[] data, int index)
    {
        var span = data.AsSpan();
        const uint TtcTag = 0x74746366; // 'ttcf'
        if (BinaryPrimitives.ReadUInt32BigEndian(span) != TtcTag) return data;

        uint fontCount = BinaryPrimitives.ReadUInt32BigEndian(span[8..]);
        if (index >= fontCount) index = 0;
        int directory = (int)BinaryPrimitives.ReadUInt32BigEndian(span[(12 + 4 * index)..]);
        int tableCount = BinaryPrimitives.ReadUInt16BigEndian(span[(directory + 4)..]);
        int headerSize = 12 + 16 * tableCount;

        var tables = new (int Offset, int Length)[tableCount];
        int total = headerSize;
        for (int i = 0; i < tableCount; i++)
        {
            int record = directory + 12 + 16 * i;
            int offset = (int)BinaryPrimitives.ReadUInt32BigEndian(span[(record + 8)..]);
            int length = (int)BinaryPrimitives.ReadUInt32BigEndian(span[(record + 12)..]);
            tables[i] = (offset, length);
            total += Align4(length);
        }

        var result = new byte[total];
        span.Slice(directory, headerSize).CopyTo(result);
        int position = headerSize;
        for (int i = 0; i < tableCount; i++)
        {
            var (offset, length) = tables[i];
            span.Slice(offset, length).CopyTo(result.AsSpan(position));
            BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(12 + 16 * i + 8), (uint)position);
            position += Align4(length);
        }
        return result;
    }

    private static int Align4(int value) => (value + 3) & ~3;
}
