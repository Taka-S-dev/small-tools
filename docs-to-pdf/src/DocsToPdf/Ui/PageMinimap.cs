namespace DocsToPdf.Ui;

/// <summary>
/// ページの並び（PageStrip）の横に置く、文書全体を表す細い帯。大きな PDF で、目当てのページへ移るためのもの。
/// 選んだページの位置に印を、章（しおり）の始まりに線を、今見えている範囲に枠を描く。
/// 押すか引きずると、その位置へ移る。指を乗せると、その位置のページ番号を出す。
/// 全ページが一度に見えているときは何も描かない（幅は空けたままにする。出し入れで並びが変わり、行き来しないように）。
/// </summary>
internal sealed class PageMinimap : Control
{
    private static readonly IReadOnlySet<int> NoPages = new HashSet<int>();

    private readonly PageStrip _strip;
    private readonly ToolTip _tip = new() { InitialDelay = 0, ReshowDelay = 0 };
    private IReadOnlySet<int> _selected = NoPages;
    private IReadOnlyList<int> _chapters = [];
    private int _tipPage = -1;
    private bool _dragging;

    public PageMinimap(PageStrip strip)
    {
        _strip = strip;
        DoubleBuffered = true;
        ResizeRedraw = true;
        Dock = DockStyle.Right;
        Width = 20;
        Cursor = Cursors.Hand;
        AccessibleName = "文書全体の中の位置";
        BackColor = Theme.PreviewBack;
        _strip.ViewChanged += Invalidate;
    }

    /// <summary>印を付ける、選ばれたページ。</summary>
    public void SetSelection(IReadOnlySet<int> selected)
    {
        _selected = selected;
        Invalidate();
    }

    /// <summary>線を引く、章（一番上の階層のしおり）の始まりのページ。</summary>
    public void SetChapters(IReadOnlyList<int> pages)
    {
        _chapters = pages;
        Invalidate();
    }

    /// <summary>スクロールしないと見えないページがあるときだけ、帯を使う。</summary>
    private bool Needed => _strip.PageCount > 0 && _strip.ContentHeight > _strip.ClientSize.Height + 1;

    /// <summary>全体の中の位置 → 帯の上の位置の倍率。</summary>
    private float ContentScale => ClientSize.Height / (float)Math.Max(1, _strip.ContentHeight);

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (!Needed) return;
        var g = e.Graphics;
        float scale = ContentScale;
        int inset = LogicalToDeviceUnits(3);

        // 帯の地: 文字の色をうっすら重ねて、ページの並びの背景と見分ける
        using (var track = new SolidBrush(Color.FromArgb(28, ForeColor))) g.FillRectangle(track, 0, 0, Width, Height);

        // 章の始まり: 帯の幅いっぱいの細い線
        using (var chapter = new Pen(SystemColors.GrayText))
        {
            foreach (int page in _chapters)
            {
                var bounds = _strip.PageBoundsOf(page);
                if (bounds.IsEmpty) continue;
                int y = (int)(bounds.Top * scale);
                g.DrawLine(chapter, 0, y, Width, y);
            }
        }

        // 選んだページ: 小さくても見えるよう、最低 2 ピクセルの高さで塗る
        using (var mark = new SolidBrush(SystemColors.Highlight))
        {
            foreach (int page in _selected)
            {
                var bounds = _strip.PageBoundsOf(page);
                if (bounds.IsEmpty) continue;
                int y = (int)(bounds.Top * scale);
                int h = Math.Max(LogicalToDeviceUnits(2), (int)(bounds.Height * scale));
                g.FillRectangle(mark, inset, y, Width - inset * 2, h);
            }
        }

        // 今見えている範囲
        int top = (int)(_strip.ViewTop * scale);
        int height = Math.Max(LogicalToDeviceUnits(6), (int)(_strip.ClientSize.Height * scale));
        using (var view = new SolidBrush(Color.FromArgb(50, SystemColors.Highlight)))
            g.FillRectangle(view, 0, top, Width - 1, height);
        using (var edge = new Pen(SystemColors.Highlight))
            g.DrawRectangle(edge, 0, top, Width - 1, height - 1);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button != MouseButtons.Left || !Needed) return;
        _dragging = true;
        Capture = true;
        MoveTo(e.Y);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (!Needed) return;
        if (_dragging) MoveTo(e.Y);
        ShowPageAt(e.Y);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        _dragging = false;
        Capture = false;
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        _tip.Hide(this);
        _tipPage = -1;
    }

    /// <summary>帯の位置 y が、見えている範囲の真ん中に来るようにスクロールする。</summary>
    private void MoveTo(int y)
    {
        int content = (int)(y / ContentScale);
        _strip.ScrollToContent(content - _strip.ClientSize.Height / 2);
    }

    /// <summary>帯の位置 y にあるページの番号を、指の横に出す。</summary>
    private void ShowPageAt(int y)
    {
        int content = (int)(y / ContentScale);
        int page = -1;
        for (int i = 0; i < _strip.PageCount; i++)
        {
            var bounds = _strip.PageBoundsOf(i);
            if (bounds.IsEmpty) continue;
            if (bounds.Top > content) break;
            page = i; // 上から見て、この位置より上で始まる最後のページ
        }
        if (page < 0 || page == _tipPage) return;
        _tipPage = page;
        _tip.Show($"{page + 1} ページ", this, -LogicalToDeviceUnits(80), y);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _strip.ViewChanged -= Invalidate;
            _tip.Dispose();
        }
        base.Dispose(disposing);
    }
}
