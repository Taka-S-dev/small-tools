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
}

internal sealed record MergeOptions(bool PageNumbers, bool TableOfContents, bool DuplexBlanks, string DocumentTitle)
{
    public static readonly MergeOptions None = new(false, false, false, "");

    /// <summary>
    /// ファイルごとのしおりを付け、元のしおりをその下に入れる（結合のとき）。
    /// false なら元のしおりをそのまま一番上の階層に置く（分割のとき。1 ファイルしかないので包む必要がない）。
    /// </summary>
    public bool FileBookmarks { get; init; } = true;
}

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
            output.Outlines.Add("目次", tocPages[0], true);
        }

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

            // 元のページ番号から、結合後のページを引けるようにする（しおりの付け替えに使う）
            var added = new PdfPage?[input.PageCount];
            PdfPage? first = null;
            foreach (int index in part.Pages)
            {
                if (index < 0 || index >= input.PageCount) continue;
                var page = output.AddPage(input.Pages[index]);
                if (part.TurnOf(index) is int turn and not 0)
                    page.Rotate = (PdfInfo.Rotation(page) + turn) % 360;
                added[index] = page;
                first ??= page;
            }
            if (first is not null && options.FileBookmarks)
            {
                var node = output.Outlines.Add(part.Title, first, false);
                if (!IsRedundant(part.Outline, part.Pages[0]))
                    AddOutline(node.Outlines, part.Outline, added);
            }
            else if (first is not null)
            {
                AddOutline(output.Outlines, part.Outline, added);
            }
            progress?.Report(i + 1);
        }

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

    /// <summary>ファイルの先頭を指すだけのしおり 1 件なら、ファイル名のしおりと重なるので付けない。</summary>
    private static bool IsRedundant(IReadOnlyList<OutlineNode> outline, int firstPage) =>
        outline.Count == 0
        || (outline.Count == 1 && outline[0].Children.Count == 0 && outline[0].PageIndex == firstPage);

    /// <param name="pages">元のページ番号 → 結合後のページ。結合に含めなかったページは null。</param>
    private static void AddOutline(PdfOutlineCollection target, IReadOnlyList<OutlineNode> nodes, PdfPage?[] pages)
    {
        foreach (var node in nodes)
        {
            var page = node.PageIndex >= 0 && node.PageIndex < pages.Length ? pages[node.PageIndex] : null;
            if (page is null)
            {
                // 行き先のページが結合に含まれない項目は付けない（子は残せるものだけ同じ階層に付ける）
                AddOutline(target, node.Children, pages);
                continue;
            }
            var added = target.Add(node.Title, page, false);
            AddOutline(added.Outlines, node.Children, pages);
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
