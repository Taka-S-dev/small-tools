using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>
/// 1 ページを大きく見せるウィンドウ。← → でページを送り、結合に使うかどうかと向きもここで変えられる。
/// 変更は呼び出し元に渡し、結果はファイルの状態（FileEntry）から読み直して表示する。
/// </summary>
internal sealed class PageViewer : Form
{
    private readonly FileEntry _entry;
    private readonly bool _editable;
    private readonly Action<int> _toggle;
    private readonly Action<int, int> _rotate;
    private readonly Action<int, IWin32Window>? _addBookmark;
    private readonly PdfPageRenderer _renderer = new();
    private readonly System.Windows.Forms.Timer _settle = new() { Interval = 150 };

    private readonly Button _prev = new() { Text = "◀ 前へ", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0, 0, 4, 0) };
    private readonly Button _next = new() { Text = "次へ ▶", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0, 0, 12, 0) };
    private readonly Label _position = new() { AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 16, 0) };
    private readonly CheckBox _include = new() { Text = "このページを結合する", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 16, 0) };
    private readonly Button _rotateLeft = new() { Text = "左に回転", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0, 0, 4, 0) };
    private readonly Button _rotateRight = new() { Text = "右に回転", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0, 0, 12, 0) };
    private readonly Button _bookmark = new() { Text = "しおりを付ける...", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0, 0, 12, 0) };
    private readonly Label _keys = new() { AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SystemColors.GrayText };
    private readonly DoubleBufferedPanel _canvas = new() { Dock = DockStyle.Fill, BackColor = Theme.ViewerBack };

    private int _page;
    private Bitmap? _image;
    private (int Page, int Turn, Size Size) _imageKey = (-1, 0, Size.Empty);
    private int _generation;
    private bool _updating;
    private Font? _veilFont;

    /// <param name="addBookmark">(ページ, 名前を聞くダイアログの親) → そのページにしおりを足す。null ならしおりのボタンを出さない。</param>
    public PageViewer(FileEntry entry, int page, bool editable, Action<int> toggle, Action<int, int> rotate,
        Action<int, IWin32Window>? addBookmark = null)
    {
        _entry = entry;
        _page = page;
        _editable = editable;
        _toggle = toggle;
        _rotate = rotate;
        _addBookmark = editable ? addBookmark : null;

        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = entry.FileName;
        ShowIcon = false;
        ShowInTaskbar = false;
        KeyPreview = true;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        MinimumSize = new Size(520, 420);
        ClientSize = new Size(900, 900); // 実際の大きさは OnLoad で画面に合わせて決める

        var bar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Padding = new Padding(10, 8, 10, 8) };
        bar.Controls.AddRange([_prev, _next, _position]);
        if (editable) bar.Controls.AddRange([_include, _rotateLeft, _rotateRight]);
        if (_addBookmark is not null) bar.Controls.Add(_bookmark);
        bar.Controls.Add(_keys);
        _keys.Text = editable
            ? "← → でページ送り・Space で結合の切り替え・R で回転" + (_addBookmark is not null ? "・B でしおり" : "") + "・Esc で閉じる"
            : "← → でページ送り・Esc で閉じる";

        Controls.Add(_canvas);
        Controls.Add(bar);

        _prev.Click += (_, _) => Go(_page - 1);
        _next.Click += (_, _) => Go(_page + 1);
        _rotateLeft.Click += (_, _) => Rotate(-90);
        _rotateRight.Click += (_, _) => Rotate(90);
        _bookmark.Click += (_, _) => AddBookmark();
        _include.CheckedChanged += (_, _) =>
        {
            if (_updating) return;
            _toggle(_page);
            ShowState(); // 外せなかったとき（最後の 1 ページ）は、チェックを元に戻す
        };
        _canvas.Paint += OnCanvasPaint;
        _canvas.Resize += (_, _) =>
        {
            _canvas.Invalidate();
            _settle.Stop();
            _settle.Start();
        };
        _canvas.MouseWheel += (_, e) => Go(_page + (e.Delta > 0 ? -1 : 1));
        _settle.Tick += (_, _) =>
        {
            _settle.Stop();
            Render();
        };
        ResumeLayout(false);
        PerformLayout();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        // 画面の作業領域（実際のピクセル）から決めるので、表示倍率の調整が済んだあとに設定する
        var area = Screen.FromControl(Owner ?? this).WorkingArea;
        Size = new Size(Math.Max(MinimumSize.Width, (int)(area.Width * 0.6)), Math.Max(MinimumSize.Height, (int)(area.Height * 0.88)));
        StartPosition = FormStartPosition.Manual;
        CenterToParent();
        // 画面からはみ出さないようにする
        Location = new Point(Math.Clamp(Left, area.Left, Math.Max(area.Left, area.Right - Width)),
            Math.Clamp(Top, area.Top, Math.Max(area.Top, area.Bottom - Height)));
    }

    private int PageCount => _entry.Converted?.PageCount ?? 0;
    private int Turn => _entry.Rotations.TryGetValue(_page, out int turn) ? turn : 0;

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        ShowState();
        _canvas.Focus();
    }

    private void Go(int page)
    {
        page = Math.Clamp(page, 0, Math.Max(0, PageCount - 1));
        if (page == _page) return;
        _page = page;
        ShowState();
    }

    private void AddBookmark()
    {
        if (_addBookmark is null) return;
        _addBookmark(_page, this);
        ShowState();
    }

    private void Rotate(int degrees)
    {
        if (!_editable) return;
        _rotate(_page, degrees);
        ShowState();
    }

    /// <summary>ページの位置・結合に使うか・向きを、ファイルの状態から読んで表示し直す。</summary>
    private void ShowState()
    {
        if (_entry.Converted is null || PageCount == 0)
        {
            Close();
            return;
        }
        _page = Math.Clamp(_page, 0, PageCount - 1);
        bool included = _entry.SelectedPages.Contains(_page);

        _updating = true;
        _include.Checked = included;
        _updating = false;

        int turn = Turn;
        _position.Text = $"{_page + 1} / {PageCount} ページ" + (turn != 0 ? $"（{turn}° 回転）" : "") + (included ? "" : "・結合しない")
            + (PreviewPane.MarkOf(_entry, _page) is { } mark ? $"・しおり「{mark}」" : "");
        _position.ForeColor = included ? ForeColor : Theme.Danger;
        _prev.Enabled = _page > 0;
        _next.Enabled = _page < PageCount - 1;
        Render();
    }

    /// <summary>キャンバスに余白を残して収まる、ページの表示位置。</summary>
    private Rectangle FitRect()
    {
        if (_entry.Converted is not { } pdf || _page >= pdf.PageCount) return Rectangle.Empty;
        var size = PageGlyphs.Turned(pdf.PageSizes[_page], Turn);
        int margin = LogicalToDeviceUnits(16);
        var area = Rectangle.Inflate(_canvas.ClientRectangle, -margin, -margin);
        if (area.Width < 16 || area.Height < 16) return Rectangle.Empty;
        double scale = Math.Min(area.Width / Math.Max(1, size.Width), area.Height / Math.Max(1, size.Height));
        int width = Math.Max(1, (int)(size.Width * scale)), height = Math.Max(1, (int)(size.Height * scale));
        return new Rectangle(area.X + (area.Width - width) / 2, area.Y + (area.Height - height) / 2, width, height);
    }

    private async void Render()
    {
        if (_entry.Converted is not { } pdf) return;
        var rect = FitRect();
        if (rect.IsEmpty) return;
        var key = (_page, Turn, rect.Size);
        if (key == _imageKey && _image is not null) return;

        int generation = ++_generation;
        Bitmap? image = null;
        try
        {
            image = await _renderer.RenderAsync(pdf.PdfPath, key.Item1, rect.Width, rect.Height, key.Item2);
        }
        catch
        {
            // 描けないときは「表示できません」と出す
        }
        if (IsDisposed || generation != _generation)
        {
            image?.Dispose();
            return;
        }
        _image?.Dispose();
        _image = image;
        _imageKey = key;
        _canvas.Invalidate();
    }

    private void OnCanvasPaint(object? sender, PaintEventArgs e)
    {
        var rect = FitRect();
        if (rect.IsEmpty) return;
        var g = e.Graphics;
        g.FillRectangle(Brushes.White, rect);

        // 描き直しが届くまでは、前の画像を伸ばして見せる（同じページで向きが同じときだけ）
        if (_image is not null && _imageKey.Page == _page && _imageKey.Turn == Turn)
        {
            g.InterpolationMode = _imageKey.Size == rect.Size
                ? System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor
                : System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
            g.DrawImage(_image, rect);
        }
        else
        {
            TextRenderer.DrawText(g, _image is null && _imageKey.Page == _page ? "表示できません" : "読み込んでいます...",
                Font, rect, Color.Gray, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter); // 白いページの上
        }

        if (!_entry.SelectedPages.Contains(_page))
        {
            using var veil = new SolidBrush(Theme.Veil(Theme.PreviewBack, 160));
            g.FillRectangle(veil, rect);
            _veilFont ??= new Font(Font.FontFamily, Font.Size * 1.6f, FontStyle.Bold);
            TextRenderer.DrawText(g, "結合しないページ", _veilFont, rect, Theme.Danger,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Escape:
                Close();
                return true;
            case Keys.Left or Keys.Up or Keys.PageUp:
                Go(_page - 1);
                return true;
            case Keys.Right or Keys.Down or Keys.PageDown:
                Go(_page + 1);
                return true;
            case Keys.Home:
                Go(0);
                return true;
            case Keys.End:
                Go(PageCount - 1);
                return true;
            case Keys.Space when _editable:
                _include.Checked = !_include.Checked;
                return true;
            case Keys.R when _editable:
                Rotate(90);
                return true;
            case Keys.Shift | Keys.R when _editable:
                Rotate(-90);
                return true;
            case Keys.B when _addBookmark is not null:
                AddBookmark();
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _settle.Dispose();
            _image?.Dispose();
            _veilFont?.Dispose();
            _renderer.Reset();
        }
        base.Dispose(disposing);
    }

    private sealed class DoubleBufferedPanel : Panel
    {
        public DoubleBufferedPanel()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            TabStop = true;
        }
    }
}
