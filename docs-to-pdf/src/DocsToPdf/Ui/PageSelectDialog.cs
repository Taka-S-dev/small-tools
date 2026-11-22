using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>
/// 1 ファイルの全ページを並べ、選んだページだけで 1 つの PDF を保存する（「1-3 と 10-12 だけ印刷したい」とき）。
/// ページを押して選び、Shift を押しながら押すと、前に押したページからそこまでをまとめて選ぶ。番号を欄に書いてもよい。
/// 一覧の結合の設定（外したページ）は変えない。回したページは回した向きで保存する。ページの番号は元の PDF の番号。
/// </summary>
internal sealed class PageSelectDialog : Form
{
    private readonly FileEntry _entry;
    private readonly ConvertedPdf _pdf;
    private readonly SortedSet<int> _selected = [];
    private int _anchor = -1;
    private bool _syncing;
    private bool _running;

    private readonly TextBox _pagesBox = new() { Dock = DockStyle.Fill, PlaceholderText = "例: 1-3,10-12" };
    private readonly Button _all = ExcelSheetsDialog.DialogButton("すべて選ぶ");
    private readonly Button _none = ExcelSheetsDialog.DialogButton("選択を外す");
    private readonly Label _hint = new()
    {
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(0, 4, 0, 6),
        Text = "ページを押すと選びます（もう一度押すと外す）。Shift を押しながら押すと、前に押したページからそこまでをまとめて選びます。",
    };
    // 抜き出すページを探す画面なので、何枚も並べて見渡せる大きさにする
    private readonly PageStrip _pages = new() { Dock = DockStyle.Fill, TargetPageWidth = 140 };
    private readonly Label _count = new() { AutoSize = true, Anchor = AnchorStyles.Left };
    private readonly Button _save = ExcelSheetsDialog.DialogButton("保存...");
    private readonly Button _cancel = ExcelSheetsDialog.DialogButton("キャンセル", DialogResult.Cancel);

    public PageSelectDialog(FileEntry entry)
    {
        _entry = entry;
        _pdf = entry.Converted ?? throw new ArgumentException("変換が済んでいないファイルからは選べません。", nameof(entry));

        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = $"{entry.FileName} からページを選んで保存";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        ShowInTaskbar = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(900, 640);
        MinimumSize = new Size(560, 420);

        var top = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 4, Margin = Padding.Empty };
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        top.Controls.Add(new Label { Text = "選ぶページ", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        top.Controls.Add(_pagesBox, 1, 0);
        top.Controls.Add(_all, 2, 0);
        top.Controls.Add(_none, 3, 0);

        var frame = new Panel { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, Margin = Padding.Empty };
        frame.Controls.Add(_pages);

        var bottom = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 3, Margin = new Padding(0, 10, 0, 0) };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.Controls.Add(_count, 0, 0);
        bottom.Controls.Add(_save, 1, 0);
        bottom.Controls.Add(_cancel, 2, 0);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(12) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(top);
        root.Controls.Add(_hint);
        root.Controls.Add(frame);
        root.Controls.Add(bottom);
        Controls.Add(root);

        // 画面の読み上げや自動操作で見分けられるよう、入力欄に名前を付ける
        (_pagesBox.Name, _pagesBox.AccessibleName) = ("Pages", "選ぶページ");

        AcceptButton = _save;
        CancelButton = _cancel;

        _pages.ShowDocument(_pdf, Enumerable.Range(0, _pdf.PageCount), entry.Rotations, interactive: false, hint: "");
        _pages.SetSelection(_selected);
        _pages.PageClicked += Pick;
        _pages.PageZoomRequested += page =>
        {
            using var viewer = new PageViewer(_entry, page, editable: false, toggle: _ => { }, rotate: (_, _) => { });
            viewer.ShowDialog(this);
        };
        _pagesBox.TextChanged += (_, _) => ReadBox();
        _all.Click += (_, _) => SetSelection(Enumerable.Range(0, _pdf.PageCount));
        _none.Click += (_, _) => SetSelection([]);
        _save.Click += async (_, _) => await SaveAsync();
        ResumeLayout(false);
        PerformLayout();
        ShowSelection(writeBox: true);
    }

    /// <summary>保存した PDF のパス（保存したあと）。</summary>
    public string? Written { get; private set; }

    /// <summary>保存した PDF のページ数。</summary>
    public int WrittenPages { get; private set; }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // 保存している途中では閉じない（書きかけのファイルが残らないよう、終わるまで待ってもらう）
        if (_running) e.Cancel = true;
        base.OnFormClosing(e);
    }

    /// <summary>ページが押された。Shift なら前に押したページからここまでを選び、そうでなければ選ぶ・外すを切り替える。</summary>
    private void Pick(int page)
    {
        if (_running) return;
        if ((ModifierKeys & Keys.Shift) != 0 && _anchor >= 0)
        {
            int from = Math.Min(_anchor, page), to = Math.Max(_anchor, page);
            for (int i = from; i <= to; i++) _selected.Add(i);
        }
        else if (!_selected.Remove(page))
        {
            _selected.Add(page);
        }
        _anchor = page;
        ShowSelection(writeBox: true);
    }

    private void SetSelection(IEnumerable<int> pages)
    {
        _selected.Clear();
        _selected.UnionWith(pages);
        _anchor = -1;
        ShowSelection(writeBox: true);
    }

    /// <summary>欄に書かれた番号を読み、選択に反映する。読めない間は選択を変えずに、その旨を出す。</summary>
    private void ReadBox()
    {
        if (_syncing) return;
        string text = _pagesBox.Text;
        if (PageRange.Normalize(text).Length == 0)
        {
            _selected.Clear();
        }
        else if (PageRange.IsValid(text))
        {
            _selected.Clear();
            _selected.UnionWith(PageRange.Resolve(text, _pdf.PageCount));
        }
        else
        {
            ShowCount("「1-3,10-12」のように、ページの番号と範囲を「,」で区切って書いてください", problem: true);
            return;
        }
        ShowSelection(writeBox: false);
    }

    /// <summary>選んだページを、ページの印・欄・枚数・保存ボタンに反映する。</summary>
    private void ShowSelection(bool writeBox)
    {
        if (writeBox)
        {
            _syncing = true;
            // 全ページのときも空にせず「1-N」と書く（空は「何も選んでいない」の意味にするため）
            _pagesBox.Text = _selected.Count == _pdf.PageCount ? $"1-{_pdf.PageCount}" : PageRange.Format(_selected, _pdf.PageCount);
            _syncing = false;
        }
        _pages.SetSelection(_selected);
        ShowCount(_selected.Count == 0 ? $"全 {_pdf.PageCount} ページから、保存するページを選んでください" : $"全 {_pdf.PageCount} ページ中 {_selected.Count} ページを選択中", problem: false);
        _save.Enabled = _selected.Count > 0 && !_running;
        _save.Text = _selected.Count > 0 ? $"選んだ {_selected.Count} ページを保存..." : "保存...";
    }

    private void ShowCount(string text, bool problem)
    {
        _count.Text = text;
        _count.ForeColor = problem ? Theme.Danger : SystemColors.ControlText;
        if (problem) _save.Enabled = false;
    }

    /// <summary>
    /// 画面の自動操作で確かめるとき用（Debug ビルドだけ）。保存先を尋ねるダイアログは自動操作しにくいので、環境変数で渡せる（MainForm と同じ）。
    /// Release ビルドでは、いつも null（保存先を尋ねる）。
    /// </summary>
    private static string? TestSavePath()
    {
#if DEBUG
        if (Environment.GetEnvironmentVariable("DOCSTOPDF_TEST_SAVE_AS") is { Length: > 0 } forced) return Path.GetFullPath(forced);
#endif
        return null;
    }

    private async Task SaveAsync()
    {
        if (_selected.Count == 0) return;
        string range = PageRange.Format(_selected, _pdf.PageCount);
        string name = PdfSplitter.SafeFileName($"{_entry.DefaultTitle}_{(range.Length > 0 ? range : "全ページ")}");
        using var dialog = new SaveFileDialog
        {
            Title = "選んだページの保存先",
            Filter = "PDF ファイル|*.pdf",
            DefaultExt = "pdf",
            AddExtension = true,
            OverwritePrompt = true,
            InitialDirectory = Path.GetDirectoryName(_entry.Path) ?? "",
            FileName = name + ".pdf",
        };
        string? path = TestSavePath();
        while (path is null && dialog.ShowDialog(this) == DialogResult.OK)
        {
            string chosen = Path.GetFullPath(dialog.FileName);
            if (!string.Equals(chosen, _entry.Path, StringComparison.OrdinalIgnoreCase))
            {
                path = chosen;
                break;
            }
            MessageBox.Show(this, "元のファイルと同じ名前では保存できません。別の名前を指定してください。",
                MainForm.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        if (path is null) return;

        SetRunning(true);
        try
        {
            var part = new MergePart(_pdf.PdfPath, _entry.Title, _selected.ToList(), _pdf.Outline, _entry.Rotations);
            int pages = await Task.Run(() => IndividualPdf.Write(part, path, pageNumbers: false));
            (Written, WrittenPages) = (path, pages);
            SetRunning(false);
            DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            SetRunning(false);
            TaskDialog.ShowDialog(this, new TaskDialogPage
            {
                Caption = MainForm.AppTitle,
                Heading = "保存できませんでした",
                Text = ex.Message,
                Icon = TaskDialogIcon.Error,
            });
        }
    }

    private void SetRunning(bool running)
    {
        _running = running;
        UseWaitCursor = running;
        foreach (Control control in new Control[] { _pagesBox, _all, _none, _cancel }) control.Enabled = !running;
        if (running)
        {
            _save.Enabled = false;
            _save.Text = "保存しています...";
        }
        else
        {
            ShowSelection(writeBox: false);
        }
    }
}
