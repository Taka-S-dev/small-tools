using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>
/// 一覧の 1 ファイルを、複数の PDF に分けて書き出す。
/// 分け方を変えるたびに、できあがるファイルの名前とページをその場で一覧に出す。
/// プレビューで結合から外したページは含めず、回したページは回した向きで書き出す。
/// ページの番号は、外したページを除いた「見えている順」で数える。
/// </summary>
internal sealed class SplitDialog : Form
{
    private readonly FileEntry _entry;
    private readonly ConvertedPdf _pdf;
    private readonly IReadOnlyList<int> _used;
    private readonly IReadOnlyList<OutlineNode> _bookmarks;
    private readonly string _baseName;
    /// <summary>しおりごとに分けたときのファイル数（しおりが無ければ 0）。</summary>
    private readonly int _bookmarkFiles;

    private readonly RadioButton _byEvery = new() { Text = "ページ数ごと", AutoSize = true };
    private readonly NumericUpDown _every = new() { Minimum = 1, Maximum = 10_000, Value = 1, Width = 72 };
    private readonly Label _everyUnit = new() { Text = "ページずつ", AutoSize = true, Anchor = AnchorStyles.Left };
    private readonly RadioButton _byBookmarks = new() { AutoSize = true };
    private readonly RadioButton _byRanges = new() { Text = "範囲を指定", AutoSize = true };
    private readonly TextBox _ranges = new() { Dock = DockStyle.Fill, PlaceholderText = "例: 1-3; 4-8; 9-" };
    private readonly Label _rangesHint = new() { AutoSize = true, ForeColor = SystemColors.GrayText, Text = "「;」で区切った範囲ごとに 1 ファイル。1 つの範囲の中では「1-3,5」のように書けます" };
    private readonly ListView _preview = new()
    {
        View = View.Details,
        FullRowSelect = true,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
        Dock = DockStyle.Fill,
        MultiSelect = false,
    };
    private readonly Label _problem = new() { AutoSize = true, ForeColor = Theme.Danger, Margin = new Padding(0, 4, 0, 0) };
    private readonly TextBox _folder = new() { Dock = DockStyle.Fill };
    private readonly Button _browse = ExcelSheetsDialog.DialogButton("参照...");
    private readonly CheckBox _reveal = new() { Text = "終わったらフォルダを開く", AutoSize = true, Checked = true };
    private readonly Button _split = ExcelSheetsDialog.DialogButton("分割");
    private readonly Button _cancel = ExcelSheetsDialog.DialogButton("キャンセル", DialogResult.Cancel);

    private IReadOnlyList<PdfSplitter.Piece> _pieces = [];
    private IReadOnlyList<string> _paths = [];
    private bool _running;

    public SplitDialog(FileEntry entry)
    {
        _entry = entry;
        _pdf = entry.Converted ?? throw new ArgumentException("変換が済んでいないファイルは分けられません。", nameof(entry));
        _used = entry.SelectedPages;
        _bookmarks = PdfSplitter.Restrict(_pdf.Outline, _used);
        _bookmarkFiles = PdfSplitter.ByBookmarks(_bookmarks, _used.Count).Count;
        _baseName = PdfSplitter.SafeFileName(entry.Title) is { Length: > 0 } safe ? safe : entry.DefaultTitle;

        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = $"{entry.FileName} を分割";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(560, 520);
        MinimumSize = new Size(460, 440);

        _byBookmarks.Text = _bookmarkFiles > 0 ? $"しおりごと（{_bookmarkFiles} ファイル）" : "しおりごと（しおりがありません）";
        _byBookmarks.Enabled = _bookmarkFiles > 0;

        _preview.Columns.Add("できるファイル");
        _preview.Columns.Add("ページ");
        _preview.Columns.Add("枚数", -2, HorizontalAlignment.Right);
        _preview.Resize += (_, _) => FitColumns();

        var everyRow = Row(_byEvery, _every, _everyUnit);
        var rangesRow = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Margin = Padding.Empty };
        rangesRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        rangesRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        rangesRow.Controls.Add(_byRanges, 0, 0);
        rangesRow.Controls.Add(_ranges, 1, 0);
        _rangesHint.Margin = new Padding(20, 0, 0, 6);

        var folderRow = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 3, Margin = new Padding(0, 8, 0, 0) };
        folderRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        folderRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        folderRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        folderRow.Controls.Add(new Label { Text = "保存先", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        folderRow.Controls.Add(_folder, 1, 0);
        folderRow.Controls.Add(_browse, 2, 0);

        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 10, 0, 0) };
        buttons.Controls.AddRange([_cancel, _split]);

        string scope = _used.Count == _pdf.PageCount
            ? $"全 {_pdf.PageCount} ページ"
            : $"結合に使う {_used.Count} ページ（外した {_pdf.PageCount - _used.Count} ページは含めません。番号はこの {_used.Count} ページで数えます）";
        var header = new Label { Text = $"{entry.FileName}　{scope}", AutoSize = true, MaximumSize = new Size(520, 0), Margin = new Padding(0, 0, 0, 8) };

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(12) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (var control in new Control[] { header, everyRow, _byBookmarks, rangesRow, _rangesHint })
        {
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(control);
        }
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.Controls.Add(_preview);
        foreach (var control in new Control[] { _problem, folderRow, _reveal, buttons })
        {
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(control);
        }
        Controls.Add(root);

        // 画面の読み上げや自動操作で見分けられるよう、入力欄に名前を付ける
        (_every.Name, _every.AccessibleName) = ("Every", "何ページずつ分けるか");
        (_ranges.Name, _ranges.AccessibleName) = ("Ranges", "分ける範囲");
        (_folder.Name, _folder.AccessibleName) = ("Folder", "保存先");
        (_preview.Name, _preview.AccessibleName) = ("Pieces", "できるファイル");

        AcceptButton = _split;
        CancelButton = _cancel;
        _folder.Text = Path.GetDirectoryName(entry.Path) ?? "";

        // 初めの分け方: しおりが 2 つ以上あれば章ごと、なければ 1 ページずつ
        if (_bookmarkFiles >= 2) _byBookmarks.Checked = true;
        else _byEvery.Checked = true;

        // 3 つのラジオボタンは、入力欄と並べるため別々の入れ物に置いている。
        // WinForms は同じ入れ物の中でしか排他にしないので、選ばれたら他を外す
        var modes = new[] { _byEvery, _byBookmarks, _byRanges };
        foreach (var radio in modes)
        {
            radio.CheckedChanged += (_, _) =>
            {
                if (!radio.Checked) return;
                foreach (var other in modes.Where(m => m != radio)) other.Checked = false;
                UpdatePlan();
            };
        }
        _every.ValueChanged += (_, _) => { _byEvery.Checked = true; UpdatePlan(); };
        _ranges.TextChanged += (_, _) => { if (_ranges.Focused) _byRanges.Checked = true; UpdatePlan(); };
        _folder.TextChanged += (_, _) => UpdatePlan();
        _browse.Click += (_, _) => ChooseFolder();
        _split.Click += async (_, _) => await SplitAsync();
        ResumeLayout(false);
        PerformLayout();
        UpdatePlan();
    }

    /// <summary>書き出したファイル（分割したあと）。</summary>
    public IReadOnlyList<string> Written { get; private set; } = [];

    public bool RevealAfterwards => _reveal.Checked;

    private static FlowLayoutPanel Row(params Control[] controls)
    {
        var row = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = Padding.Empty };
        foreach (var control in controls) control.Margin = new Padding(0, 3, 6, 3);
        row.Controls.AddRange(controls);
        return row;
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // 書き出している途中では閉じない（途中のファイルが残らないよう、終わるまで待ってもらう）
        if (_running) e.Cancel = true;
        base.OnFormClosing(e);
    }

    /// <summary>今の指定で、できあがるファイルを計算し直して一覧に出す。</summary>
    private void UpdatePlan()
    {
        string? problem = null;
        try
        {
            var pieces = _byEvery.Checked ? PdfSplitter.Every(_used.Count, (int)_every.Value)
                : _byBookmarks.Checked ? PdfSplitter.ByBookmarks(_bookmarks, _used.Count)
                : PdfSplitter.Ranges(PdfSplitter.SplitRangeList(_ranges.Text), _used.Count);
            _pieces = PdfSplitter.ToSource(pieces, _used);
            _paths = Directory.Exists(_folder.Text) ? PdfSplitter.FileNames(_baseName, _pieces, _folder.Text) : [];
            if (!Directory.Exists(_folder.Text)) problem = "保存先のフォルダが見つかりません。";
            else if (_paths.Contains(_entry.Path, StringComparer.OrdinalIgnoreCase)) problem = "元のファイルと同じ名前になるファイルがあります。保存先を変えてください。";
            ShowPieces(pieces);
        }
        catch (FormatException ex)
        {
            _pieces = [];
            _paths = [];
            problem = ex.Message;
            ShowPieces([]);
        }
        _problem.Text = problem ?? "";
        _problem.Visible = problem is not null;
        _split.Enabled = problem is null && _pieces.Count > 0 && !_running;
        _split.Text = _pieces.Count > 0 ? $"{_pieces.Count} ファイルに分割" : "分割";
    }

    /// <param name="positions">見えている順（1 始まりで表示する）で数えた分け方。</param>
    private void ShowPieces(IReadOnlyList<PdfSplitter.Piece> positions)
    {
        _preview.BeginUpdate();
        _preview.Items.Clear();
        var names = _paths.Count == positions.Count ? _paths : PdfSplitter.FileNames(_baseName, positions, "");
        for (int i = 0; i < positions.Count; i++)
        {
            var item = new ListViewItem(Path.GetFileName(names[i]));
            item.SubItems.Add(PageRange.Format(positions[i].Pages, int.MaxValue) is { Length: > 0 } range ? range : "すべて");
            item.SubItems.Add(positions[i].Pages.Count.ToString());
            if (_paths.Count == positions.Count && File.Exists(_paths[i]))
            {
                item.ForeColor = Theme.Danger;
                item.ToolTipText = "同じ名前のファイルがあります（分割すると上書きします）";
            }
            _preview.Items.Add(item);
        }
        _preview.ShowItemToolTips = true;
        _preview.EndUpdate();
    }

    /// <summary>ファイル名の列に残りの幅をすべて配る（ページと枚数の列は文字に合わせた幅）。</summary>
    private void FitColumns()
    {
        if (_preview.Columns.Count < 3) return;
        int pages = LogicalToDeviceUnits(110), count = LogicalToDeviceUnits(56);
        int name = Math.Max(LogicalToDeviceUnits(120), _preview.ClientSize.Width - pages - count - 1);
        _preview.Columns[0].Width = name;
        _preview.Columns[1].Width = pages;
        _preview.Columns[2].Width = count;
    }

    private void ChooseFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "分割したファイルの保存先",
            UseDescriptionForTitle = true,
            InitialDirectory = Directory.Exists(_folder.Text) ? _folder.Text : "",
        };
        if (dialog.ShowDialog(this) == DialogResult.OK) _folder.Text = dialog.SelectedPath;
    }

    private async Task SplitAsync()
    {
        if (_pieces.Count == 0 || _paths.Count != _pieces.Count) return;

        var existing = _paths.Where(File.Exists).ToList();
        if (existing.Count > 0)
        {
            var overwrite = new TaskDialogButton("上書きする");
            var page = new TaskDialogPage
            {
                Caption = MainForm.AppTitle,
                Heading = $"同じ名前のファイルが {existing.Count} 件あります",
                Text = string.Join("\r\n", existing.Take(5).Select(Path.GetFileName)) + (existing.Count > 5 ? $"\r\nほか {existing.Count - 5} 件" : ""),
                Icon = TaskDialogIcon.Warning,
                Buttons = { overwrite, TaskDialogButton.Cancel },
                DefaultButton = TaskDialogButton.Cancel,
            };
            if (TaskDialog.ShowDialog(this, page) != overwrite) return;
        }

        SetRunning(true);
        try
        {
            var pieces = _pieces;
            var paths = _paths;
            var rotations = _entry.Rotations;
            await Task.Run(() => PdfSplitter.Split(_pdf.PdfPath, _pdf.Outline, _entry.Title, pieces, paths, rotations: rotations));
            Written = paths;
            SetRunning(false);
            DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            SetRunning(false);
            TaskDialog.ShowDialog(this, new TaskDialogPage
            {
                Caption = MainForm.AppTitle,
                Heading = "分割できませんでした",
                Text = ex.Message,
                Icon = TaskDialogIcon.Error,
            });
        }
    }

    private void SetRunning(bool running)
    {
        _running = running;
        UseWaitCursor = running;
        foreach (Control control in new Control[] { _byEvery, _every, _byBookmarks, _byRanges, _ranges, _folder, _browse, _cancel })
            control.Enabled = !running && (control != _byBookmarks || _bookmarkFiles > 0);
        _split.Text = running ? "分割しています..." : _split.Text;
        if (!running) UpdatePlan();
        else _split.Enabled = false;
    }
}
