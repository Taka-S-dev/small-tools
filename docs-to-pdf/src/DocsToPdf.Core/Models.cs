namespace DocsToPdf.Core;

internal enum FileKind { Word, Excel, PowerPoint, Pdf, Image }

internal enum EntryState { Queued, Converting, Ready, Failed, NeedsPassword }

/// <summary>Excel の表示されているシート 1 枚。Pages は出力されるページ数（対象外や空のシートは 0）。</summary>
internal sealed record SheetInfo(string Name, int Pages, bool Included);

/// <summary>しおり 1 件。PageIndex はその PDF の中での 0 始まりのページ位置。</summary>
internal sealed record OutlineNode(string Title, int PageIndex, IReadOnlyList<OutlineNode> Children);

/// <summary>画面に表示したときのページの大きさ（ポイント。回転を反映済み）。</summary>
internal readonly record struct PageSize(double Width, double Height);

/// <summary>元のファイルが変わったかを見分けるための更新日時とサイズ。</summary>
internal readonly record struct SourceStamp(DateTime LastWriteUtc, long Length)
{
    public static SourceStamp? Of(string path)
    {
        try
        {
            var info = new FileInfo(path);
            return info.Exists ? new SourceStamp(info.LastWriteTimeUtc, info.Length) : null;
        }
        catch
        {
            return null;
        }
    }
}

/// <summary>変換を終えて作業フォルダに置いた PDF と、その中身の情報。</summary>
internal sealed record ConvertedPdf(
    string PdfPath,
    IReadOnlyList<PageSize> PageSizes,
    IReadOnlyList<OutlineNode> Outline,
    IReadOnlyList<SheetInfo>? Sheets)
{
    public int PageCount => PageSizes.Count;
}

/// <summary>変換の依頼。変換スレッドへ渡すので、行の状態を写した変わらない値にしてある。</summary>
internal sealed record ConvertRequest(
    string Path,
    FileKind Kind,
    IReadOnlyList<string>? Sheets,
    bool FitWidth,
    string? Password);

/// <summary>変換の結果。成功なら Pdf、失敗なら Error が入る。</summary>
internal sealed record ConvertOutcome(ConvertedPdf? Pdf, SourceStamp? Stamp, string? Error, bool NeedsPassword)
{
    public static ConvertOutcome Success(ConvertedPdf pdf, SourceStamp? stamp) => new(pdf, stamp, null, false);
    public static ConvertOutcome Failure(string error, SourceStamp? stamp, bool needsPassword = false) => new(null, stamp, error, needsPassword);
}

/// <summary>結合リストの 1 行。UI スレッドだけが読み書きする。</summary>
internal sealed class FileEntry
{
    private static readonly Dictionary<string, FileKind> Extensions = new(StringComparer.OrdinalIgnoreCase)
    {
        [".doc"] = FileKind.Word,
        [".docx"] = FileKind.Word,
        [".docm"] = FileKind.Word,
        [".rtf"] = FileKind.Word,
        [".xls"] = FileKind.Excel,
        [".xlsx"] = FileKind.Excel,
        [".xlsm"] = FileKind.Excel,
        [".xlsb"] = FileKind.Excel,
        [".ppt"] = FileKind.PowerPoint,
        [".pptx"] = FileKind.PowerPoint,
        [".pptm"] = FileKind.PowerPoint,
        [".pps"] = FileKind.PowerPoint,
        [".ppsx"] = FileKind.PowerPoint,
        [".pdf"] = FileKind.Pdf,
        [".jpg"] = FileKind.Image,
        [".jpeg"] = FileKind.Image,
        [".png"] = FileKind.Image,
        [".bmp"] = FileKind.Image,
        [".gif"] = FileKind.Image,
        [".tif"] = FileKind.Image,
        [".tiff"] = FileKind.Image,
    };

    public static IEnumerable<string> SupportedExtensions => Extensions.Keys;

    private (string Range, int Count, IReadOnlyList<int> Pages)? _selectedCache;

    public FileEntry(string path, FileKind kind)
    {
        Path = path;
        Kind = kind;
        Title = DefaultTitle;
    }

    public string Path { get; }
    public FileKind Kind { get; }
    /// <summary>フォルダを入れて見つけたファイルなら、そのフォルダ。しおりをフォルダごとにまとめるのに使う。</summary>
    public string? Root { get; init; }
    public string FileName => System.IO.Path.GetFileName(Path);
    public string DefaultTitle => System.IO.Path.GetFileNameWithoutExtension(Path);

    // ── 利用者が決めるもの ──
    /// <summary>しおりと目次に出す名前。</summary>
    public string Title { get; set; }
    /// <summary>使うページの指定（"1-3,5"）。空ならすべて。</summary>
    public string PageRange { get; set; } = "";
    /// <summary>出力する Excel のシート名。null ならすべて。</summary>
    public IReadOnlyList<string>? Sheets { get; set; }
    /// <summary>Excel のすべての列を 1 ページの幅に収める。</summary>
    public bool FitWidth { get; set; }
    public string? Password { get; set; }
    /// <summary>利用者が回したページ。0 始まりのページ番号 → 右回りの角度（90 / 180 / 270）。</summary>
    public IReadOnlyDictionary<int, int> Rotations { get; set; } = NoRotations;

    public static readonly IReadOnlyDictionary<int, int> NoRotations = new Dictionary<int, int>();

    /// <summary>結合に使うページのうち、回したページの数。</summary>
    public int RotatedCount => Rotations.Count == 0 ? 0 : SelectedPages.Count(Rotations.ContainsKey);

    // ── 変換の状態 ──
    public EntryState State { get; set; } = EntryState.Queued;
    public string? Error { get; set; }
    public ConvertedPdf? Converted { get; set; }
    public SourceStamp? Stamp { get; set; }
    /// <summary>変換を頼み直すたびに増やす。古い依頼の結果を捨てるのに使う。</summary>
    public int Version { get; set; }
    public CancellationTokenSource? Pending { get; set; }

    /// <summary>結合に使うページ（0 始まり）。変換が済んでいなければ空。</summary>
    public IReadOnlyList<int> SelectedPages
    {
        get
        {
            if (Converted is null) return [];
            int count = Converted.PageCount;
            if (_selectedCache is { } c && c.Range == PageRange && c.Count == count) return c.Pages;
            var pages = Core.PageRange.Resolve(PageRange, count);
            _selectedCache = (PageRange, count, pages);
            return pages;
        }
    }

    public bool IsBusy => State is EntryState.Queued or EntryState.Converting;
    /// <summary>結合に使える状態か（変換が済み、使うページが 1 枚以上ある）。</summary>
    public bool IsUsable => State == EntryState.Ready && SelectedPages.Count > 0;

    public string KindLabel => Kind switch
    {
        FileKind.Word => "Word",
        FileKind.Excel => "Excel",
        FileKind.PowerPoint => "PowerPoint",
        FileKind.Image => "画像",
        _ => "PDF",
    };

    public static bool TryGetKind(string path, out FileKind kind)
    {
        kind = default;
        // "~$" で始まるのは Office が開いている間だけ作るロックファイル
        if (System.IO.Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            return false;
        return Extensions.TryGetValue(System.IO.Path.GetExtension(path), out kind);
    }
}
