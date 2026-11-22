using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Core;

/// <summary>
/// 結合する 1 ファイル分。Pages は使うページ（0 始まり、ページ順）。
/// Rotations は、さらに右へ回すページ（0 始まりのページ番号 → 90 / 180 / 270）。
/// </summary>
internal sealed record MergePart(string PdfPath, string Title, IReadOnlyList<int> Pages, IReadOnlyList<OutlineNode> Outline,
    IReadOnlyDictionary<int, int>? Rotations = null)
{
    public int TurnOf(int page) => Rotations is not null && Rotations.TryGetValue(page, out int turn) ? turn : 0;

    /// <summary>このファイルのしおりを入れる、フォルダの段（上から順）。空なら一番上の階層に置く。</summary>
    public IReadOnlyList<BookmarkFolder> Folders { get; init; } = [];

    /// <summary>
    /// このファイルだけ、元のしおりを何段まで付けるか（PowerPoint のスライドだけ外す、など）。
    /// null なら MergeOptions.OutlineDepth に従う。すべて付けるときは AllLevels、付けないときは 0。
    /// </summary>
    public int? OutlineDepth { get; init; }

    public const int AllLevels = int.MaxValue;
}

internal sealed record MergeOptions(bool PageNumbers, bool TableOfContents, bool DuplexBlanks, string DocumentTitle)
{
    public static readonly MergeOptions None = new(false, false, false, "");

    /// <summary>
    /// ファイルごとのしおりを付け、元のしおりをその下に入れる（結合のとき）。
    /// false なら元のしおりをそのまま一番上の階層に置く（分割のとき。1 ファイルしかないので包む必要がない）。
    /// </summary>
    public bool FileBookmarks { get; init; } = true;

    /// <summary>元のファイルのしおり（Word の見出しなど）を何段まで付けるか。null ならすべて、0 なら付けない。</summary>
    public int? OutlineDepth { get; init; }
}

internal enum BookmarkKind { Contents, Folder, File, Original }

/// <summary>結合した PDF に付けるしおり 1 件。Page は結合後のページ（0 始まり）。</summary>
/// <param name="Part">ファイルのしおりなら、そのファイルの番号（parts の何番目か）。ほかは -1。</param>
/// <param name="FolderKey">フォルダの段なら、そのフォルダの Key。ほかは null。</param>
internal sealed record Bookmark(string Title, int Page, BookmarkKind Kind, IReadOnlyList<Bookmark> Children,
    int Part = -1, string? FolderKey = null);

internal static class PdfMerger
{
    private const double A4Width = 595.276, A4Height = 841.89;
    private const double TocMargin = 56, TocTitleHeight = 52, TocRowHeight = 23;
    private static readonly int TocRowsPerPage = (int)((A4Height - TocMargin * 2 - TocTitleHeight) / TocRowHeight);

    /// <summary>目次を付けたときの目次のページ数。</summary>
    public static int TocPageCount(int partCount) => (partCount + TocRowsPerPage - 1) / TocRowsPerPage;

    /// <summary>
    /// 結合後の総ページ数と、各ファイルの開始ページ（0 始まり）を求める。
    /// 両面印刷用の白紙は、各ファイルが奇数ページ（紙の表）から始まるように前へ挟む。
    /// </summary>
    public static (int TotalPages, int[] Starts, bool[] BlankBefore) Plan(IReadOnlyList<int> pageCounts, MergeOptions options)
    {
        var starts = new int[pageCounts.Count];
        var blankBefore = new bool[pageCounts.Count];
        int cursor = options.TableOfContents ? TocPageCount(pageCounts.Count) : 0;
        for (int i = 0; i < pageCounts.Count; i++)
        {
            if (options.DuplexBlanks && cursor % 2 == 1)
            {
                blankBefore[i] = true;
                cursor++;
            }
            starts[i] = cursor;
            cursor += pageCounts[i];
        }
        return (cursor, starts, blankBefore);
    }

    /// <summary>
    /// 結合した PDF に付くしおりを求める（作る前に確かめるため。結合でも同じ決まりで付ける）。
    /// 目次 → フォルダの段 → ファイル → 元のしおり、の順に入れ子にする。
    /// 続けて並んだファイルが同じフォルダなら 1 つの段にまとめ、間に別のフォルダが挟まれば段を分ける。
    /// フォルダの中身が同じ名前のファイル 1 つだけなら（見積/見積.xlsx）、段を作らずファイルだけを置く。
    /// 比べるのはフォルダの元の名前（付け直した名前で比べると、名前を変えただけで段が消え、木から戻せなくなる）。
    /// </summary>
    public static IReadOnlyList<Bookmark> Bookmarks(IReadOnlyList<MergePart> parts, MergeOptions options)
    {
        var (_, starts, _) = Plan(parts.Select(p => p.Pages.Count).ToList(), options);
        return Bookmarks(parts, options, (i, k) => starts[i] + k);
    }

    /// <param name="pageAt">(ファイルの番号, そのファイルで使うページの何枚目) → 結合後のページ。結合に入らなければ null。</param>
    private static IReadOnlyList<Bookmark> Bookmarks(IReadOnlyList<MergePart> parts, MergeOptions options, Func<int, int, int?> pageAt)
    {
        var top = new List<Bookmark>();
        if (options.TableOfContents && TocPageCount(parts.Count) > 0)
            top.Add(new Bookmark("目次", 0, BookmarkKind.Contents, []));

        var open = new List<(string Key, List<Bookmark> Children)>(); // 今開いているフォルダの段（上から順）
        for (int i = 0; i < parts.Count; i++)
        {
            var part = parts[i];
            var positions = new Dictionary<int, int>();
            for (int k = 0; k < part.Pages.Count; k++)
                if (pageAt(i, k) is int at) positions.TryAdd(part.Pages[k], at);
            if (positions.Count == 0) continue;
            int first = positions.Values.Min();

            var originals = new List<Bookmark>();
            int depth = part.OutlineDepth ?? options.OutlineDepth ?? MergePart.AllLevels;
            AddOriginal(originals, part.Outline, positions, 1, depth);
            // 段数で削ったあとに、ファイルの先頭を指す 1 件だけが残ったなら、ファイルのしおりと重なるので付けない
            if (options.FileBookmarks && originals is [{ Children.Count: 0 } only] && only.Page == first) originals.Clear();
            if (!options.FileBookmarks)
            {
                top.AddRange(originals);
                continue;
            }

            // 前のファイルと同じところまでの段はそのまま使い、違うところから先を閉じて開き直す
            int keep = 0;
            while (keep < open.Count && keep < part.Folders.Count
                && string.Equals(open[keep].Key, part.Folders[keep].Key, StringComparison.OrdinalIgnoreCase))
                keep++;
            open.RemoveRange(keep, open.Count - keep);
            for (int d = keep; d < part.Folders.Count; d++)
            {
                var folder = part.Folders[d];
                var children = new List<Bookmark>();
                (d == 0 ? top : open[d - 1].Children).Add(new Bookmark(folder.Name, first, BookmarkKind.Folder, children, FolderKey: folder.Key));
                open.Add((folder.Key, children));
            }
            (open.Count == 0 ? top : open[^1].Children).Add(new Bookmark(part.Title, first, BookmarkKind.File, originals, i));
        }
        return Collapse(top);
    }

    /// <summary>中身が同じ名前のファイル 1 つだけのフォルダの段を、そのファイルに置き換える（「見積 › 見積」にしない）。</summary>
    private static List<Bookmark> Collapse(IReadOnlyList<Bookmark> nodes) =>
        nodes.Select(node =>
        {
            if (node.Kind != BookmarkKind.Folder) return node;
            var children = Collapse(node.Children);
            return children is [{ Kind: BookmarkKind.File } only] && node.FolderKey is { } key
                && string.Equals(only.Title, BookmarkFolders.NameOf(key), StringComparison.OrdinalIgnoreCase)
                ? only
                : node with { Children = children };
        }).ToList();

    /// <param name="positions">元のページ番号 → 結合後のページ。結合に含めなかったページは入っていない。</param>
    /// <param name="level">nodes の段（元のしおりの一番上が 1）。maxDepth を超える段は付けない。</param>
    private static void AddOriginal(List<Bookmark> target, IReadOnlyList<OutlineNode> nodes, Dictionary<int, int> positions,
        int level, int maxDepth)
    {
        if (level > maxDepth) return;
        foreach (var node in nodes)
        {
            var children = new List<Bookmark>();
            AddOriginal(children, node.Children, positions, level + 1, maxDepth);
            // 行き先のページが結合に含まれない項目は付けない（子は残せるものだけ同じ階層に付ける）
            if (positions.TryGetValue(node.PageIndex, out int page)) target.Add(new Bookmark(node.Title, page, BookmarkKind.Original, children));
            else target.AddRange(children);
        }
    }

    /// <summary>
    /// 順に結合し、しおりを付けて保存する。途中で失敗したり中止されたりしても、保存先は壊さない。
    /// </summary>
    /// <returns>できた PDF のページ数。</returns>
    public static int Merge(IReadOnlyList<MergePart> parts, MergeOptions options, string outputPath,
        IProgress<int>? progress = null, CancellationToken ct = default)
    {
        string dir = Path.GetDirectoryName(Path.GetFullPath(outputPath))!;
        string temp = Path.Combine(dir, $"~{Path.GetFileNameWithoutExtension(outputPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            int pageCount;
            using (var output = new PdfDocument())
            {
                Build(output, parts, options, progress, ct);
                pageCount = output.PageCount;
                ct.ThrowIfCancellationRequested();
                output.Save(temp);
            }
            ct.ThrowIfCancellationRequested();
            File.Move(temp, outputPath, overwrite: true);
            return pageCount;
        }
        finally
        {
            try { if (File.Exists(temp)) File.Delete(temp); } catch { /* 残っても害はない */ }
        }
    }

    private static void Build(PdfDocument output, IReadOnlyList<MergePart> parts, MergeOptions options,
        IProgress<int>? progress, CancellationToken ct)
    {
        output.PageMode = PdfPageMode.UseOutlines;
        output.Info.Title = options.DocumentTitle;
        output.Info.Creator = "Docs to PDF";

        var (_, starts, blankBefore) = Plan(parts.Select(p => p.Pages.Count).ToList(), options);

        var tocPages = new List<PdfPage>();
        if (options.TableOfContents)
        {
            for (int i = 0; i < TocPageCount(parts.Count); i++)
                tocPages.Add(AddBlank(output, A4Width, A4Height));
        }
        // ファイルごとの「使うページの何枚目 → 結合後のページ」。しおりは、全部のページを足してから付ける
        var placed = new List<int?[]>();

        var blanks = new HashSet<PdfPage>();
        for (int i = 0; i < parts.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var part = parts[i];
            using var input = PdfReader.Open(part.PdfPath, PdfDocumentOpenMode.Import);

            if (blankBefore[i] && part.Pages.Count > 0)
            {
                // 白紙は、続くページを回したあとの向きにそろえる
                var size = PdfInfo.Read(input).Sizes[part.Pages[0]];
                bool sideways = part.TurnOf(part.Pages[0]) is 90 or 270;
                blanks.Add(sideways ? AddBlank(output, size.Height, size.Width) : AddBlank(output, size.Width, size.Height));
            }

            var at = new int?[part.Pages.Count];
            for (int k = 0; k < part.Pages.Count; k++)
            {
                int index = part.Pages[k];
                if (index < 0 || index >= input.PageCount) continue;
                var page = output.AddPage(input.Pages[index]);
                if (part.TurnOf(index) is int turn and not 0)
                    page.Rotate = (PdfInfo.Rotation(page) + turn) % 360;
                at[k] = output.PageCount - 1;
            }
            placed.Add(at);
            progress?.Report(i + 1);
        }

        AddBookmarks(output.Outlines, Bookmarks(parts, options, (i, k) => placed[i][k]), output);

        if (options.TableOfContents)
            DrawToc(output, tocPages, parts, starts);
        if (options.PageNumbers)
            DrawPageNumbers(output, blanks);
    }

    private static PdfPage AddBlank(PdfDocument output, double width, double height)
    {
        var page = output.AddPage();
        page.Width = XUnit.FromPoint(width);
        page.Height = XUnit.FromPoint(height);
        return page;
    }

    /// <summary>しおりを PDF に書く。目次とフォルダの段は開いた状態にし、ファイルと元のしおりは閉じておく。</summary>
    private static void AddBookmarks(PdfOutlineCollection target, IReadOnlyList<Bookmark> nodes, PdfDocument output)
    {
        foreach (var node in nodes)
        {
            var added = target.Add(node.Title, output.Pages[node.Page], node.Kind is BookmarkKind.Contents or BookmarkKind.Folder);
            AddBookmarks(added.Outlines, node.Children, output);
        }
    }

    private static void DrawToc(PdfDocument output, List<PdfPage> tocPages, IReadOnlyList<MergePart> parts, int[] starts)
    {
        PdfFonts.EnsureRegistered();
        var titleFont = new XFont(PdfFonts.Family, 18);
        var rowFont = new XFont(PdfFonts.Family, 10.5);
        var leaderPen = new XPen(XColor.FromArgb(170, 170, 170), 0.6) { DashStyle = XDashStyle.Dot };
        double right = A4Width - TocMargin;

        for (int pageIndex = 0; pageIndex < tocPages.Count; pageIndex++)
        {
            var page = tocPages[pageIndex];
            using var g = XGraphics.FromPdfPage(page);
            g.DrawString(pageIndex == 0 ? "目次" : "目次（続き）", titleFont, XBrushes.Black, new XPoint(TocMargin, TocMargin + 20));

            int firstRow = pageIndex * TocRowsPerPage;
            int lastRow = Math.Min(parts.Count, firstRow + TocRowsPerPage);
            for (int row = firstRow; row < lastRow; row++)
            {
                if (parts[row].Pages.Count == 0) continue;
                double top = TocMargin + TocTitleHeight + (row - firstRow) * TocRowHeight;
                double baseline = top + 15;
                string number = (starts[row] + 1).ToString();
                double numberWidth = g.MeasureString(number, rowFont).Width;
                double maxTitleWidth = right - TocMargin - numberWidth - 24;
                string title = Fit(g, rowFont, parts[row].Title, maxTitleWidth);
                double titleWidth = g.MeasureString(title, rowFont).Width;

                g.DrawString(title, rowFont, XBrushes.Black, new XPoint(TocMargin, baseline));
                g.DrawString(number, rowFont, XBrushes.Black, new XPoint(right - numberWidth, baseline));
                double leaderStart = TocMargin + titleWidth + 8, leaderEnd = right - numberWidth - 8;
                if (leaderEnd > leaderStart)
                    g.DrawLine(leaderPen, leaderStart, baseline - 3, leaderEnd, baseline - 3);

                // 行のどこを押しても、そのファイルの先頭ページへ飛ぶ
                var area = g.Transformer.WorldToDefaultPage(new XRect(TocMargin, top, right - TocMargin, TocRowHeight));
                page.AddDocumentLink(new PdfRectangle(area), starts[row] + 1);
            }
        }
    }

    /// <summary>幅に収まらない文字列を「…」で切る。</summary>
    private static string Fit(XGraphics g, XFont font, string text, double maxWidth)
    {
        if (g.MeasureString(text, font).Width <= maxWidth) return text;
        int length = text.Length;
        while (length > 1 && g.MeasureString(text[..length] + "…", font).Width > maxWidth)
            length--;
        return text[..length] + "…";
    }

    private static void DrawPageNumbers(PdfDocument output, HashSet<PdfPage> blanks)
    {
        PdfFonts.EnsureRegistered();
        var font = new XFont(PdfFonts.Family, 9);
        var brush = new XSolidBrush(XColor.FromArgb(70, 70, 70));
        int total = output.PageCount;

        for (int i = 0; i < total; i++)
        {
            var page = output.Pages[i];
            if (blanks.Contains(page)) continue;
            using var g = OpenUpright(page, out double width, out double height);
            g.DrawString($"{i + 1} / {total}", font, brush, new XRect(0, height - 28, width, 12), XStringFormats.Center);
        }
    }

    /// <summary>
    /// 画面で見たとおりの向き・範囲で描けるようにした描画面を開く。
    /// 原点は見えている範囲の左上、width / height は見たときの幅と高さ。
    /// ページの回転（/Rotate）と、原点が 0 でない MediaBox や CropBox を吸収する。
    /// </summary>
    private static XGraphics OpenUpright(PdfPage page, out double width, out double height)
    {
        var media = page.MediaBox;
        var box = PdfInfo.VisibleBox(page);
        int rotation = PdfInfo.Rotation(page);
        double boxWidth = box.Width, boxHeight = box.Height;

        var g = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
        // PDFsharp の描画座標 (x, y) は、PDF の座標 (x, MediaBox の高さ - y) に置かれる
        // （MediaBox の原点が 0 でなくても考慮されない）。見えている範囲の左上が原点になるよう移す。
        g.TranslateTransform(box.X1, media.Height - box.Y2);
        switch (rotation)
        {
            case 90:
                g.TranslateTransform(0, boxHeight);
                g.RotateTransform(-90);
                break;
            case 180:
                g.TranslateTransform(boxWidth, boxHeight);
                g.RotateTransform(180);
                break;
            case 270:
                g.TranslateTransform(boxWidth, 0);
                g.RotateTransform(90);
                break;
        }

        bool sideways = rotation is 90 or 270;
        width = sideways ? boxHeight : boxWidth;
        height = sideways ? boxWidth : boxHeight;
        return g;
    }
}
