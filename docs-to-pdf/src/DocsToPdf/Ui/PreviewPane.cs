using System.Drawing.Drawing2D;
using System.Runtime.InteropServices.WindowsRuntime;
using DocsToPdf.Core;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace DocsToPdf.Ui;

/// <summary>PDF のページを画像にする（Windows に備わっている PDF の描画機能を使う）。UI スレッドから呼ぶこと。</summary>
internal sealed class PdfPageRenderer
{
    private string? _path;
    private PdfDocument? _document;
    private IRandomAccessStream? _stream;

    /// <param name="displayWidth">回したあとの、表示する幅（ピクセル）。</param>
    /// <param name="turn">右へ回す角度（0 / 90 / 180 / 270）。</param>
    public async Task<Bitmap> RenderAsync(string path, int pageIndex, int displayWidth, int displayHeight, int turn)
    {
        if (_path != path || _document is null)
        {
            // ファイルを開いたまま、描くページに要る部分だけを読む（数百 MB の PDF でも丸ごとメモリに載せない）。
            // 開いているのは作業フォルダの変換結果で、変換し直すと別の名前のファイルになるので、開いたままでも邪魔にならない。
            // ほかのアプリで開く・消すのは妨げないよう、読み書きと削除の共有を許す
            var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete,
                bufferSize: 1 << 16, FileOptions.Asynchronous | FileOptions.RandomAccess);
            IRandomAccessStream stream = file.AsRandomAccessStream();
            PdfDocument document;
            try
            {
                document = await PdfDocument.LoadFromStreamAsync(stream);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
            _stream?.Dispose();
            (_path, _document, _stream) = (path, document, stream);
        }

        // 横に倒すときは、元のページの幅が表示の高さになる
        bool sideways = turn is 90 or 270;
        int width = sideways ? displayHeight : displayWidth;

        using var page = _document.GetPage((uint)pageIndex);
        using var output = new InMemoryRandomAccessStream();
        await page.RenderToStreamAsync(output, new PdfPageRenderOptions { DestinationWidth = (uint)Math.Max(16, width) });
        using var decoded = new Bitmap(output.AsStream());
        var image = new Bitmap(decoded); // ストリームから切り離した写しを返す
        image.RotateFlip(turn switch
        {
            90 => RotateFlipType.Rotate90FlipNone,
            180 => RotateFlipType.Rotate180FlipNone,
            270 => RotateFlipType.Rotate270FlipNone,
            _ => RotateFlipType.RotateNoneFlipNone,
        });

        // Windows の PDF 描画は、幅を画素ではなく表示倍率を掛けた単位で受け取るため、頼んだ幅の 2 倍などで返ってくる。
        // 描く側は表示する大きさの画像を 1 対 1 で置く（間引いて縮めると細い線や文字がつぶれる）ので、
        // ここで表示するちょうどの大きさに、丁寧に縮めて（大きく返ってきた分はかえって滑らかになる）そろえる
        if (image.Width != displayWidth || image.Height != displayHeight)
        {
            var fitted = new Bitmap(Math.Max(1, displayWidth), Math.Max(1, displayHeight));
            using (var g = Graphics.FromImage(fitted))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                using var wrap = new System.Drawing.Imaging.ImageAttributes();
                wrap.SetWrapMode(WrapMode.TileFlipXY); // 縁が半透明ににじまないようにする
                g.DrawImage(image, new Rectangle(0, 0, fitted.Width, fitted.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrap);
            }
            image.Dispose();
            image = fitted;
        }
        return image;
    }

    public void Reset()
    {
        _stream?.Dispose();
        (_path, _document, _stream) = (null, null, null);
    }
}

/// <summary>ページの向きと、ページに重ねる小さなボタンの記号。</summary>
internal static class PageGlyphs
{
    // Windows 10 以降に入っている記号のフォント。なければ文字で代用する
    private const string SymbolFamily = "Segoe MDL2 Assets";
    public static readonly bool HasSymbols = FontFamily.Families.Any(f => f.Name == SymbolFamily);
    public static string Zoom => HasSymbols ? "" : "＋";
    public static string RotateRight => HasSymbols ? "" : "↻";

    public static Font SymbolFont(float size) => HasSymbols ? new Font(SymbolFamily, size) : new Font(SystemFonts.DefaultFont.FontFamily, size);

    /// <summary>回したあとのページの大きさ。</summary>
    public static PageSize Turned(PageSize size, int turn) => turn is 90 or 270 ? new PageSize(size.Height, size.Width) : size;

    public static int Normalize(int degrees) => ((degrees % 360) + 360) % 360;
}

/// <summary>
/// ページを並べて見せる。押されたページは PageClicked で知らせる。
/// マウスを乗せたページの右上に「拡大」「回転」のボタンを出し、右クリックでも同じ操作ができる。
/// </summary>
internal sealed class PageStrip : ScrollableControl
{
    private const int MaxCachedPages = 48;
    private const int MinPageWidth = 70, MaxPageWidth = 480;
    private int _targetPageWidth = 190;

    /// <summary>
    /// 1 ページをこの幅（96dpi での値）より小さくしない範囲で、横に並べる。全体を見渡したい画面では小さくする。
    /// Ctrl を押しながらホイールを回しても変えられる（大きな PDF で、多くのページを一度に見渡すため）。
    /// </summary>
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public int TargetPageWidth
    {
        get => _targetPageWidth;
        set
        {
            value = Math.Clamp(value, MinPageWidth, MaxPageWidth);
            if (value == _targetPageWidth) return;
            // 一番上に見えていたページを、大きさを変えたあとも一番上に残す
            int top = VisiblePages(0).DefaultIfEmpty(-1).First();
            _targetPageWidth = value;
            Relayout();
            if (top >= 0) ScrollToPage(top);
            else Invalidate();
            _settle.Stop();
            _settle.Start();
        }
    }

    /// <summary>見えている範囲・全体の高さ・ページの並びが変わった（ミニマップを描き直すため）。</summary>
    public event Action? ViewChanged;

    /// <summary>PDF のページ数（並べていないページも含む）。</summary>
    public int PageCount => _pages.Length;

    /// <summary>並べたページ全体の高さ（ピクセル）。</summary>
    public int ContentHeight => Math.Max(AutoScrollMinSize.Height, ClientSize.Height);

    /// <summary>今見えている範囲の上端（全体の中での位置、ピクセル）。</summary>
    public int ViewTop => -AutoScrollPosition.Y;

    /// <summary>そのページが全体の中のどこに置かれているか（並べていなければ空）。</summary>
    public Rectangle PageBoundsOf(int page) => page >= 0 && page < _pages.Length ? _pages[page] : Rectangle.Empty;

    /// <summary>全体の中の位置 y が上に来るようにスクロールする。</summary>
    public void ScrollToContent(int y)
    {
        AutoScrollPosition = new Point(0, Math.Clamp(y, 0, Math.Max(0, ContentHeight - ClientSize.Height)));
        Invalidate();
        EnsureRendered();
        ViewChanged?.Invoke();
    }

    private enum Spot { None, Page, Zoom, Rotate, Cut }

    private sealed record CachedPage(Bitmap? Image, int Width, int Turn);

    private readonly PdfPageRenderer _renderer = new();
    private readonly Dictionary<int, CachedPage> _cache = [];
    private readonly System.Windows.Forms.Timer _settle = new() { Interval = 120 };
    private readonly ContextMenuStrip _menu = new();
    private readonly ToolStripMenuItem _menuZoom = new("大きく表示");
    private readonly ToolStripMenuItem _menuRight = new("右に 90° 回転");
    private readonly ToolStripMenuItem _menuLeft = new("左に 90° 回転");
    private readonly ToolStripMenuItem _menuToggle = new();
    private readonly Font _symbolFont;

    private ConvertedPdf? _pdf;
    private HashSet<int> _included = [];
    private IReadOnlyDictionary<int, int> _rotations = FileEntry.NoRotations;
    /// <summary>並べるページ（0 始まり、並べる順）。null なら全ページ。並べないページの枠は空のまま。</summary>
    private IReadOnlyList<int>? _shown;
    private Rectangle[] _pages = [];
    private string _message = "";
    private string _hint = "";
    private int _hintHeight;
    private bool _interactive;
    /// <summary>区切りモード: ページに「ここから分ける」ボタンを出し、押されたら PageCutToggled を知らせる（分割の画面で、ページを見ながら区切るとき）。</summary>
    private bool _cutting;
    /// <summary>新しいファイルが始まるページ → 何番目のファイルか（1 始まり）。区切りの印を描くのに使う。</summary>
    private IReadOnlyDictionary<int, int> _starts = new Dictionary<int, int>();
    /// <summary>選択モードで選ばれているページ。null なら選択モードではない（ページを選んで保存する画面で使う）。</summary>
    private IReadOnlySet<int>? _selected;
    private int _hover = -1;
    private Spot _hoverSpot;
    private int _menuPage = -1;
    private int _generation;
    private bool _rendering;

    public PageStrip()
    {
        DoubleBuffered = true;
        ResizeRedraw = true;
        AutoScroll = true;
        Name = "PagePreview";
        AccessibleName = "ページのプレビュー";
        BackColor = Theme.PreviewBack;
        _symbolFont = PageGlyphs.SymbolFont(10);
        _settle.Tick += (_, _) =>
        {
            _settle.Stop();
            EnsureRendered();
        };

        _menu.Items.AddRange([_menuZoom, new ToolStripSeparator(), _menuRight, _menuLeft, new ToolStripSeparator(), _menuToggle]);
        _menuZoom.Click += (_, _) => { if (_menuPage >= 0) PageZoomRequested?.Invoke(_menuPage); };
        _menuRight.Click += (_, _) => { if (_menuPage >= 0) PageRotateRequested?.Invoke(_menuPage, 90); };
        _menuLeft.Click += (_, _) => { if (_menuPage >= 0) PageRotateRequested?.Invoke(_menuPage, -90); };
        _menuToggle.Click += (_, _) => { if (_menuPage >= 0) PageClicked?.Invoke(_menuPage); };
    }

    public event Action<int>? PageClicked;
    public event Action<int>? PageZoomRequested;
    /// <summary>ページと、回す角度（右へ 90、左へ -90）。</summary>
    public event Action<int, int>? PageRotateRequested;
    /// <summary>区切りモードで「ここから分ける」「区切りを外す」が押されたページ。</summary>
    public event Action<int>? PageCutToggled;

    /// <summary>ページの下に出す番号。null なら元の PDF でのページ番号（1 始まり）。</summary>
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public Func<int, string>? PageLabel { get; set; }

    /// <summary>区切りモードを切り替え、新しいファイルが始まるページ（→ 何番目のファイルか）を描く。</summary>
    public void SetCuts(bool cutting, IReadOnlyDictionary<int, int> starts)
    {
        _cutting = cutting;
        _starts = starts;
        Invalidate();
    }

    /// <summary>
    /// 選択モードを切り替える。選択モードでは、ページを押すと PageClicked を知らせ（選ぶ・外すは呼び出し側が決める）、
    /// 選ばれたページに枠とチェックの印を付ける。null で選択モードをやめる。
    /// </summary>
    public void SetSelection(IReadOnlySet<int>? selected)
    {
        _selected = selected;
        Invalidate();
    }

    /// <summary>そのページが上に来るようにスクロールする。</summary>
    public void ScrollToPage(int page)
    {
        if (page < 0 || page >= _pages.Length || _pages[page].IsEmpty) return;
        AutoScrollPosition = new Point(0, Math.Max(0, _pages[page].Top - Margin12));
        Invalidate();
        EnsureRendered();
        ViewChanged?.Invoke();
    }

    /// <summary>ページの代わりに、中央に文を出す。</summary>
    public void ShowMessage(string message)
    {
        if (_pdf is null && _message == message) return;
        _pdf = null;
        _message = message;
        _generation++;
        ClearCache();
        _renderer.Reset();
        Relayout();
        Invalidate();
    }

    /// <param name="interactive">ページを押して外したり、回したりできるか。</param>
    /// <param name="hint">ページの上に出す、使い方の説明。</param>
    /// <param name="shown">並べるページ（0 始まり）。null なら全ページ。分割でできる 1 ファイル分だけを見せるときに使う。</param>
    public void ShowDocument(ConvertedPdf pdf, IEnumerable<int> included, IReadOnlyDictionary<int, int> rotations,
        bool interactive, string hint, IReadOnlyList<int>? shown = null)
    {
        _included = [.. included];
        _interactive = interactive;
        bool shownChanged = !(_shown is null ? shown is null : shown is not null && _shown.SequenceEqual(shown));
        bool layoutChanged = !ReferenceEquals(_rotations, rotations) || _hint != hint || shownChanged;
        _rotations = rotations;
        _hint = hint;
        _shown = shown;
        if (shownChanged) AutoScrollPosition = Point.Empty;
        if (!ReferenceEquals(_pdf, pdf))
        {
            _pdf = pdf;
            _message = "";
            _generation++;
            _hover = -1;
            ClearCache();
            AutoScrollPosition = Point.Empty;
            layoutChanged = true;
        }
        if (layoutChanged) Relayout();
        Invalidate();
        EnsureRendered();
    }

    private int TurnOf(int page) => _rotations.TryGetValue(page, out int turn) ? turn : 0;

    private int Margin12 => LogicalToDeviceUnits(12);
    private int LabelHeight => LogicalToDeviceUnits(20);
    private int Gap => LogicalToDeviceUnits(6);
    private int ButtonSize => LogicalToDeviceUnits(26);

    private const TextFormatFlags HintFlags = TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl | TextFormatFlags.NoPadding;

    private void Relayout()
    {
        if (_pdf is null)
        {
            _pages = [];
            AutoScrollMinSize = Size.Empty;
            return;
        }

        int available = ClientSize.Width - Margin12 * 2;
        if (available < 24)
        {
            _pages = [];
            return;
        }

        // 使い方の説明は、幅に合わせて折り返して全文を出す
        _hintHeight = _hint.Length == 0 ? 0
            : TextRenderer.MeasureText(_hint, Font, new Size(available, int.MaxValue), HintFlags).Height + Gap;

        var order = _shown?.Where(i => i >= 0 && i < _pdf.PageCount).ToList() ?? Enumerable.Range(0, _pdf.PageCount).ToList();

        // 幅に余裕があれば横に並べる。1 枚が小さくなりすぎない範囲で列を増やす
        int columns = Math.Max(1, (available + Gap) / (LogicalToDeviceUnits(TargetPageWidth) + Gap));
        columns = Math.Min(columns, Math.Max(1, order.Count));
        int width = (available - Gap * (columns - 1)) / columns;

        var pages = new Rectangle[_pdf.PageCount];
        int y = Margin12 + _hintHeight;
        for (int first = 0; first < order.Count; first += columns)
        {
            int rowHeight = 0;
            for (int slot = first; slot < Math.Min(order.Count, first + columns); slot++)
            {
                int i = order[slot];
                var size = PageGlyphs.Turned(_pdf.PageSizes[i], TurnOf(i));
                int height = Math.Max(8, (int)Math.Round(width * size.Height / Math.Max(1, size.Width)));
                pages[i] = new Rectangle(Margin12 + (slot - first) * (width + Gap), y, width, height);
                rowHeight = Math.Max(rowHeight, height);
            }
            y += rowHeight + LabelHeight + Gap;
        }
        _pages = pages;
        AutoScrollMinSize = new Size(0, y + Margin12 - Gap);
        ViewChanged?.Invoke();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        Relayout();
        // 幅を変えている間は古い画像を伸ばして見せ、落ち着いてから描き直す
        _settle.Stop();
        _settle.Start();
    }

    protected override void OnScroll(ScrollEventArgs se)
    {
        base.OnScroll(se);
        Invalidate();
        EnsureRendered();
        ViewChanged?.Invoke();
    }

    protected override void WndProc(ref Message m)
    {
        const int WmMouseWheel = 0x020A, MkControl = 0x0008;
        long wParam = m.WParam.ToInt64();
        if (m.Msg == WmMouseWheel && (wParam & MkControl) != 0)
        {
            // Ctrl+ホイール: ページの大きさを 1 段ずつ変える（スクロールはしない）。
            // Ctrl が押されていたかは、ホイールを回したときの状態（メッセージに入っている）で見る
            int delta = (short)((wParam >> 16) & 0xFFFF);
            TargetPageWidth = (int)Math.Round(TargetPageWidth * (delta > 0 ? 1.25 : 0.8));
            return;
        }
        base.WndProc(ref m);
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        Invalidate();
        UpdateHover(PointToClient(MousePosition));
        EnsureRendered();
        ViewChanged?.Invoke();
    }

    private Rectangle OnScreen(Rectangle page) =>
        new(page.X + AutoScrollPosition.X, page.Y + AutoScrollPosition.Y, page.Width, page.Height);

    private IEnumerable<int> VisiblePages(int extra)
    {
        int top = -AutoScrollPosition.Y - extra, bottom = -AutoScrollPosition.Y + ClientSize.Height + extra;
        for (int i = 0; i < _pages.Length; i++)
        {
            if (_pages[i].IsEmpty) continue; // 並べていないページ
            if (_pages[i].Bottom + LabelHeight < top) continue;
            if (_pages[i].Top > bottom) continue;
            yield return i;
        }
    }

    /// <summary>
    /// 区切りモードでページの左下に出すボタンの文字。1 つ目のファイルの先頭は外せない区切りなので出さない。
    /// ページを押しただけで区切りが変わらないよう、区切りはこのボタンでだけ変える。
    /// </summary>
    private string? CutTextOf(int page) =>
        !_cutting ? null
        : _starts.TryGetValue(page, out int file) ? (file == 1 ? null : "区切りを外す")
        : "ここから分ける";

    private Rectangle CutButtonOf(Rectangle page, string text)
    {
        int inset = LogicalToDeviceUnits(4), pad = LogicalToDeviceUnits(8);
        int width = Math.Min(page.Width - inset * 2, TextRenderer.MeasureText(text, Font).Width + pad * 2);
        int height = ButtonSize;
        return new Rectangle(page.Left + inset, page.Bottom - inset - height, Math.Max(0, width), height);
    }

    /// <summary>ページの右上に重ねるボタンの位置。右から「拡大」「回転」の順。</summary>
    private (Rectangle Zoom, Rectangle Rotate) ButtonsOf(Rectangle page)
    {
        int size = Math.Min(ButtonSize, Math.Max(12, page.Width / 3));
        int inset = LogicalToDeviceUnits(4);
        var zoom = new Rectangle(page.Right - inset - size, page.Top + inset, size, size);
        var rotate = new Rectangle(zoom.Left - inset - size, zoom.Top, size, size);
        return (zoom, rotate);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        if (_pdf is null)
        {
            TextRenderer.DrawText(g, _message, Font, ClientRectangle, SystemColors.GrayText,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
            return;
        }

        if (_hintHeight > 0)
        {
            var hint = new Rectangle(Margin12, Margin12 + AutoScrollPosition.Y, ClientSize.Width - Margin12 * 2, _hintHeight - Gap);
            TextRenderer.DrawText(g, _hint, Font, hint, SystemColors.GrayText, HintFlags);
        }

        using var border = new Pen(Theme.PageBorder);
        using var hoverBorder = new Pen(SystemColors.Highlight, LogicalToDeviceUnits(2));
        using var veil = new SolidBrush(Theme.Veil(BackColor, 185));

        foreach (int i in VisiblePages(0))
        {
            var rect = OnScreen(_pages[i]);
            bool included = _included.Contains(i);

            g.FillRectangle(Brushes.White, rect);
            if (_cache.TryGetValue(i, out var cached) && cached.Image is not null)
            {
                g.InterpolationMode = cached.Width == rect.Width && cached.Turn == TurnOf(i)
                    ? InterpolationMode.NearestNeighbor : InterpolationMode.Bilinear;
                g.DrawImage(cached.Image, rect);
            }
            else if (cached is { Image: null })
            {
                TextRenderer.DrawText(g, "表示できません", Font, rect, Color.Gray, // 白いページの上に出すので、モードによらず灰色
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }

            if (!included)
            {
                g.FillRectangle(veil, rect);
                using var cross = new Pen(Theme.Cross, LogicalToDeviceUnits(1));
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.DrawLine(cross, rect.Left, rect.Top, rect.Right, rect.Bottom);
                g.DrawLine(cross, rect.Right, rect.Top, rect.Left, rect.Bottom);
                g.SmoothingMode = SmoothingMode.Default;
            }

            if (i == _hover)
            {
                if (_interactive || _selected is not null) g.DrawRectangle(hoverBorder, rect);
                else g.DrawRectangle(border, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
                var (zoom, rotate) = ButtonsOf(rect);
                DrawButton(g, zoom, PageGlyphs.Zoom, _hoverSpot == Spot.Zoom);
                if (_interactive) DrawButton(g, rotate, PageGlyphs.RotateRight, _hoverSpot == Spot.Rotate);
                if (CutTextOf(i) is { } cutText) DrawTextButton(g, CutButtonOf(rect, cutText), cutText, _hoverSpot == Spot.Cut);
            }
            else
            {
                g.DrawRectangle(border, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
            }

            // 選択モードで選ばれたページは、太い枠と左上のチェックの印で見せる
            bool picked = _selected?.Contains(i) == true;
            if (picked) DrawPicked(g, rect);

            // 新しいファイルが始まるページは、左に縦線を引き、下に「ここから N つ目」と出す
            bool starts = _starts.TryGetValue(i, out int fileNumber);
            if (starts)
            {
                int bar = LogicalToDeviceUnits(3);
                using var accent = new SolidBrush(SystemColors.Highlight);
                g.FillRectangle(accent, rect.X - bar - LogicalToDeviceUnits(1), rect.Y - LogicalToDeviceUnits(2), bar, rect.Height + LogicalToDeviceUnits(4));
            }

            var label = new Rectangle(rect.X, rect.Bottom, rect.Width, LabelHeight);
            int turn = TurnOf(i);
            string number = PageLabel?.Invoke(i) ?? $"{i + 1}";
            string text = number + (picked ? "・選択" : "") + (starts ? $"・ここから {fileNumber} つ目" : "") + (turn != 0 ? $"・{turn}° 回転" : "") + (included ? "" : "（結合しない）");
            TextRenderer.DrawText(g, text, Font, label, !included ? Theme.Danger : starts || picked ? Theme.Link : ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }

    /// <summary>選ばれたページの印: ページを囲む太い枠と、左上の丸いチェック。</summary>
    private void DrawPicked(Graphics g, Rectangle page)
    {
        int thick = LogicalToDeviceUnits(3);
        using var frame = new Pen(SystemColors.Highlight, thick);
        g.DrawRectangle(frame, Rectangle.Inflate(page, thick / 2 + 1, thick / 2 + 1));

        int size = Math.Min(ButtonSize, Math.Max(12, page.Width / 4));
        var badge = new Rectangle(page.Left + LogicalToDeviceUnits(4), page.Top + LogicalToDeviceUnits(4), size, size);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var fill = new SolidBrush(SystemColors.Highlight)) g.FillEllipse(fill, badge);
        using (var tick = new Pen(SystemColors.HighlightText, Math.Max(2, size / 8)))
        {
            g.DrawLines(tick, [
                new Point(badge.Left + size * 27 / 100, badge.Top + size * 52 / 100),
                new Point(badge.Left + size * 43 / 100, badge.Top + size * 68 / 100),
                new Point(badge.Left + size * 73 / 100, badge.Top + size * 35 / 100),
            ]);
        }
        g.SmoothingMode = SmoothingMode.Default;
    }

    /// <summary>文字のボタン（「ここから分ける」など）。記号のボタンと同じ見た目で、文字だけ通常のフォントで描く。</summary>
    private void DrawTextButton(Graphics g, Rectangle rect, string text, bool hot)
    {
        using var fill = new SolidBrush(hot ? SystemColors.Highlight : Theme.OverlayButton);
        using var edge = new Pen(hot ? SystemColors.Highlight : Theme.OverlayButtonEdge);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var path = RoundedRect(rect, LogicalToDeviceUnits(4)))
        {
            g.FillPath(fill, path);
            g.DrawPath(edge, path);
        }
        g.SmoothingMode = SmoothingMode.Default;
        TextRenderer.DrawText(g, text, Font, rect, hot ? SystemColors.HighlightText : Theme.OverlayButtonText,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private void DrawButton(Graphics g, Rectangle rect, string glyph, bool hot)
    {
        using var fill = new SolidBrush(hot ? SystemColors.Highlight : Theme.OverlayButton);
        using var edge = new Pen(hot ? SystemColors.Highlight : Theme.OverlayButtonEdge);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var path = RoundedRect(rect, LogicalToDeviceUnits(4)))
        {
            g.FillPath(fill, path);
            g.DrawPath(edge, path);
        }
        g.SmoothingMode = SmoothingMode.Default;
        TextRenderer.DrawText(g, glyph, _symbolFont, rect, hot ? SystemColors.HighlightText : Theme.OverlayButtonText,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
    }

    private static GraphicsPath RoundedRect(Rectangle r, int radius)
    {
        int d = radius * 2;
        var path = new GraphicsPath();
        path.AddArc(r.Left, r.Top, d, d, 180, 90);
        path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private async void EnsureRendered()
    {
        if (_rendering || _pdf is null || !IsHandleCreated) return;
        _rendering = true;
        try
        {
            while (!IsDisposed && _pdf is { } pdf)
            {
                int generation = _generation;
                int index = NextToRender();
                if (index < 0) break;

                var rect = _pages[index];
                int turn = TurnOf(index);
                Bitmap? image = null;
                try
                {
                    image = await _renderer.RenderAsync(pdf.PdfPath, index, rect.Width, rect.Height, turn);
                }
                catch
                {
                    // 描けないページは「表示できません」と出す
                }

                if (IsDisposed || generation != _generation)
                {
                    image?.Dispose();
                    continue;
                }
                if (_cache.Remove(index, out var old)) old.Image?.Dispose();
                _cache[index] = new CachedPage(image, rect.Width, turn);
                TrimCache();
                if (index < _pages.Length)
                    Invalidate(Rectangle.Inflate(OnScreen(_pages[index]), 2, 2));
            }
        }
        finally
        {
            _rendering = false;
        }
    }

    /// <summary>見えているページを優先し、次にその前後を用意する。</summary>
    private int NextToRender()
    {
        foreach (int extra in new[] { 0, ClientSize.Height })
        {
            foreach (int i in VisiblePages(extra))
            {
                if (!_cache.TryGetValue(i, out var cached)) return i;
                if (cached.Turn != TurnOf(i)) return i;
                if (cached.Image is not null && Math.Abs(cached.Width - _pages[i].Width) > 1) return i;
            }
        }
        return -1;
    }

    private void TrimCache()
    {
        if (_cache.Count <= MaxCachedPages) return;
        var visible = VisiblePages(ClientSize.Height).ToHashSet();
        int center = visible.Count > 0 ? (int)visible.Average() : 0;
        foreach (int key in _cache.Keys.Where(k => !visible.Contains(k)).OrderByDescending(k => Math.Abs(k - center)).ToList())
        {
            if (_cache.Count <= MaxCachedPages) break;
            if (_cache.Remove(key, out var old)) old.Image?.Dispose();
        }
    }

    private void ClearCache()
    {
        foreach (var cached in _cache.Values) cached.Image?.Dispose();
        _cache.Clear();
    }

    private (int Page, Spot Spot) HitTest(Point point)
    {
        if (_pdf is null) return (-1, Spot.None);
        foreach (int i in VisiblePages(0))
        {
            var rect = OnScreen(_pages[i]);
            var (zoom, rotate) = ButtonsOf(rect);
            if (zoom.Contains(point)) return (i, Spot.Zoom);
            if (CutTextOf(i) is { } cutText && CutButtonOf(rect, cutText).Contains(point)) return (i, Spot.Cut);
            if (_interactive && rotate.Contains(point)) return (i, Spot.Rotate);
            rect.Height += LabelHeight;
            if (rect.Contains(point)) return (i, Spot.Page);
        }
        return (-1, Spot.None);
    }

    private void UpdateHover(Point point)
    {
        var (page, spot) = HitTest(point);
        Cursor = spot is Spot.Zoom or Spot.Rotate or Spot.Cut || (spot == Spot.Page && (_interactive || _selected is not null)) ? Cursors.Hand : Cursors.Default;
        if (page == _hover && spot == _hoverSpot) return;
        _hover = page;
        _hoverSpot = spot;
        Invalidate();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        UpdateHover(e.Location);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        UpdateHover(new Point(-1, -1));
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        base.OnMouseClick(e);
        var (page, spot) = HitTest(e.Location);
        if (page < 0) return;

        if (e.Button == MouseButtons.Right)
        {
            _menuPage = page;
            // 見るだけのときは、使えない操作を出さない（拡大だけにする）
            foreach (ToolStripItem item in _menu.Items) item.Visible = _interactive || item == _menuZoom;
            _menuToggle.Text = _included.Contains(page) ? "このページを結合しない" : "このページを結合する";
            _menu.Show(this, e.Location);
            return;
        }
        if (e.Button != MouseButtons.Left) return;

        switch (spot)
        {
            case Spot.Zoom:
                PageZoomRequested?.Invoke(page);
                break;
            case Spot.Rotate:
                PageRotateRequested?.Invoke(page, 90);
                break;
            case Spot.Page when _interactive || _selected is not null:
                PageClicked?.Invoke(page);
                break;
            case Spot.Cut:
                PageCutToggled?.Invoke(page);
                break;
            case Spot.Page:
                PageZoomRequested?.Invoke(page); // 作成中など変更できないときは、押すと大きく見せる
                break;
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _settle.Dispose();
            _menu.Dispose();
            _symbolFont.Dispose();
            ClearCache();
            _renderer.Reset();
        }
        base.Dispose(disposing);
    }
}

/// <summary>選んだファイルの変換結果を見せる右側の領域。見出しに、そのファイルに対してできる操作を並べる。</summary>
internal sealed class PreviewPane : UserControl
{
    private readonly Label _title = new() { AutoEllipsis = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 2) };
    private readonly Label _info = new() { AutoEllipsis = true, Dock = DockStyle.Fill, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 0, 0, 2) };
    private readonly FlowLayoutPanel _actions = new() { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = Padding.Empty };
    private readonly PageStrip _strip = new() { Dock = DockStyle.Fill };
    private FileEntry? _entry;

    public PreviewPane()
    {
        _title.Font = new Font(Font, FontStyle.Bold);

        var header = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 1, Padding = new Padding(8, 4, 4, 6) };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        header.Controls.Add(_title);
        header.Controls.Add(_info);
        header.Controls.Add(_actions);

        Controls.Add(_strip);
        Controls.Add(header);

        _strip.PageClicked += index =>
        {
            if (_entry is not null) PageToggled?.Invoke(_entry, index);
        };
        _strip.PageZoomRequested += index =>
        {
            if (_entry is not null) PageZoomRequested?.Invoke(_entry, index);
        };
        _strip.PageRotateRequested += (index, degrees) =>
        {
            if (_entry is not null) PageRotateRequested?.Invoke(_entry, index, degrees);
        };
    }

    /// <summary>
    /// 見出しの高さは文字の高さ（すでに表示倍率どおり）から決めるので、フォームの倍率の調整が済んだあと、
    /// ウィンドウができた時点で設定する。先に設定すると、倍率の分だけさらに拡大されてしまう。
    /// </summary>
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        _title.Height = _title.PreferredHeight;
        _info.Height = _info.PreferredHeight;
    }

    public event Action<FileEntry, int>? PageToggled;
    public event Action<FileEntry, int>? PageZoomRequested;
    public event Action<FileEntry, int, int>? PageRotateRequested;
    public event Action<FileEntry>? ChooseSheetsRequested;
    public event Action<FileEntry>? SplitRequested;
    public event Action<FileEntry>? ExtractRequested;
    public event Action<FileEntry>? PasswordRequested;
    public event Action<FileEntry>? RetryRequested;
    public event Action<FileEntry>? UseAllPagesRequested;
    public event Action<FileEntry>? ResetSheetsRequested;

    public FileEntry? Entry => _entry;

    /// <param name="entry">1 件だけ選ばれているときの、その行。</param>
    /// <param name="locked">作成中で、変更を受け付けないとき。</param>
    public void Display(FileEntry? entry, int selectedCount, int totalCount, bool locked)
    {
        _entry = entry;
        _actions.SuspendLayout();
        foreach (Control old in _actions.Controls.Cast<Control>().ToList()) old.Dispose();
        _actions.Controls.Clear();

        if (entry is null)
        {
            _title.Text = selectedCount > 1 ? $"{selectedCount} 件を選択中" : "プレビュー";
            _info.Text = "";
            _strip.ShowMessage(totalCount == 0 ? ""
                : selectedCount > 1 ? "1 件だけ選ぶと、変換結果が出ます"
                : "ファイルを選ぶと、ここに変換結果が出ます");
            _actions.ResumeLayout();
            return;
        }

        _title.Text = entry.FileName;
        switch (entry.State)
        {
            case EntryState.Queued:
                _info.Text = entry.KindLabel;
                _strip.ShowMessage("変換の順番を待っています...");
                break;

            case EntryState.Converting:
                _info.Text = entry.KindLabel;
                _strip.ShowMessage("PDF に変換しています...");
                break;

            case EntryState.NeedsPassword:
                _info.Text = entry.KindLabel;
                _strip.ShowMessage((entry.Error ?? "パスワードが必要です。") + "\r\n\r\n上の「パスワードを入力」から入力すると、変換できます。");
                if (!locked) AddAction("パスワードを入力...", () => PasswordRequested?.Invoke(entry));
                break;

            case EntryState.Failed:
                _info.Text = entry.KindLabel;
                _strip.ShowMessage($"PDF にできませんでした。\r\n\r\n{entry.Error}");
                if (!locked)
                {
                    AddAction("もう一度変換", () => RetryRequested?.Invoke(entry));
                    if (entry.Sheets is not null)
                        AddAction("シートの選択を解除", () => ResetSheetsRequested?.Invoke(entry));
                }
                break;

            default:
                var pdf = entry.Converted!;
                int used = entry.SelectedPages.Count;
                int rotated = entry.RotatedCount;
                _info.Text = (used == pdf.PageCount
                        ? $"{entry.KindLabel}・{pdf.PageCount} ページ"
                        : $"{entry.KindLabel}・{pdf.PageCount} ページ中 {used} ページを結合")
                    + (rotated > 0 ? $"・{rotated} ページを回転" : "");
                string hint = locked
                    ? "作成中は変更できません。押すと拡大表示。"
                    : "押すと結合から外す／戻す。右上のボタンで拡大・回転。";
                _strip.ShowDocument(pdf, entry.SelectedPages, entry.Rotations, interactive: !locked, hint);
                if (!locked)
                {
                    AddAction("ページを選んで保存...", () => ExtractRequested?.Invoke(entry));
                    AddAction("分割...", () => SplitRequested?.Invoke(entry));
                    if (entry.Kind == FileKind.Excel && pdf.Sheets is not null)
                        AddAction("シートと幅の設定...", () => ChooseSheetsRequested?.Invoke(entry));
                    if (used != pdf.PageCount)
                        AddAction("すべてのページを使う", () => UseAllPagesRequested?.Invoke(entry));
                }
                break;
        }
        _actions.ResumeLayout();
    }

    private void AddAction(string text, Action onClick)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(4, 1, 4, 1),
            Margin = new Padding(0, 2, 6, 0),
        };
        // 操作の結果この見出しが作り直され、押されたボタン自身が破棄されるので、処理はイベントを抜けてから行う
        button.Click += (_, _) => BeginInvoke(onClick);
        _actions.Controls.Add(button);
    }
}
