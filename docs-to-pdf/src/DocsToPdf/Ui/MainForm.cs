using System.Diagnostics;
using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>
/// メインの画面。左にファイルの一覧、右に選んだファイルのプレビューを置く。
/// ファイルは一覧に入った時点で裏で PDF にしておき、「結合して PDF を作成」では結合だけを行う。
/// このファイルは画面の組み立てとウィンドウまわり。一覧の操作は MainForm.List.cs、作成は MainForm.Create.cs。
/// </summary>
internal sealed partial class MainForm : Form
{
    public const string AppTitle = "Docs to PDF";

    // 結合の順の番号（ColOrder）は表示では先頭に出すが、アイコンをファイル名の横に残すため、項目としては最後の列に置く
    private const int ColName = 0, ColTitle = 1, ColPages = 2, ColRange = 3, ColState = 4, ColOrder = 5;

    private readonly AppSettings _settings = AppSettings.Load();
    private readonly StartupOptions _startup;
    private readonly TempWorkspace _workspace = TempWorkspace.Create();
    private readonly ConversionService _service;
    private readonly TaskbarProgress _taskbar = new();

    private readonly BufferedListView _list = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        HideSelection = false,
        MultiSelect = true,
        AllowDrop = true,
        ShowItemToolTips = true,
        HeaderStyle = ColumnHeaderStyle.Nonclickable,
        BorderStyle = BorderStyle.FixedSingle,
    };
    private readonly ColumnHeader _colName = new() { Text = "ファイル名" };
    private readonly ColumnHeader _colTitle = new() { Text = "しおり名" };
    private readonly ColumnHeader _colPages = new() { Text = "ページ", TextAlign = HorizontalAlignment.Right };
    private readonly ColumnHeader _colRange = new() { Text = "ページ指定" };
    private readonly ColumnHeader _colState = new() { Text = "状態" };
    private readonly ColumnHeader _colOrder = new() { Text = "順", TextAlign = HorizontalAlignment.Right };
    private readonly ImageList _icons = new() { ColorDepth = ColorDepth.Depth32Bit };

    private readonly Label _emptyHint = new()
    {
        Dock = DockStyle.Fill,
        TextAlign = ContentAlignment.MiddleCenter,
        ForeColor = SystemColors.GrayText,
        AllowDrop = true,
        Cursor = Cursors.Hand,
        Text = "Word・Excel・PowerPoint・PDF・画像のファイルやフォルダを\r\nここにドロップ\r\n\r\n"
             + "クリックで選ぶ、コピーしたファイルを Ctrl+V で貼り付ける、でも追加できます",
    };
    private readonly Panel _insertMark = new() { Visible = false, BackColor = SystemColors.Highlight };
    private readonly PreviewPane _preview = new() { Dock = DockStyle.Fill };
    private readonly SplitContainer _split = new() { Dock = DockStyle.Fill, SplitterWidth = 6 };

    private readonly Button _addFiles = ToolButton("ファイルを追加...");
    private readonly Button _addFolder = ToolButton("フォルダを追加...");
    private readonly Button _moveUp = ToolButton("▲ 上へ");
    private readonly Button _moveDown = ToolButton("▼ 下へ");
    private readonly Button _sortByName = ToolButton("名前順");
    private readonly Button _remove = ToolButton("削除");
    private readonly Button _clear = ToolButton("すべて削除");
    private readonly Button _undo = ToolButton("元に戻す");

    private readonly CheckBox _optPageNumbers = Option("ページ番号を入れる");
    private readonly CheckBox _optToc = Option("目次ページを付ける");
    private readonly CheckBox _optDuplex = Option("両面印刷用に白紙を挟む");
    private readonly CheckBox _optOpen = Option("作成後に開く");

    private readonly LinkLabel _status = new() { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, LinkArea = new LinkArea(0, 0) };
    private readonly Label _summary = new() { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ProgressBar _progress = new() { Dock = DockStyle.Fill, Visible = false };
    private readonly Button _run = new() { Text = "結合して PDF を作成...", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    private readonly Button _runEach = new() { Text = "個別に PDF を保存...", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    private readonly Button _cancel = new() { Text = "中止", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Enabled = false };
    private readonly ToolTip _toolTip = new();

    /// <summary>画面が出る前に、あとから起動された自分が送ってきたファイル。画面が出たら一覧に加える。</summary>
    private readonly List<string> _earlyPaths = [];
    private bool _shown;

    private CancellationTokenSource? _cts;
    private bool _busy;
    private bool _closeWhenStopped;
    private string? _lastOutput;

    public MainForm(StartupOptions startup)
    {
        _startup = startup;
        _service = new ConversionService(_workspace.FolderPath);
        // スクリプトから使ったときは、PDF を作れたときだけ成功（0）にする。途中で閉じられたら失敗のまま
        ExitCode = startup.IsBatch ? 1 : 0;

        // 大きさや余白は 96dpi での値で書き、画面の表示倍率に合わせて拡大させる。
        // 倍率の調整は ResumeLayout まで待たせないと、あとから設定した大きさが拡大されない（デザイナーが生成するコードと同じ形）
        SuspendLayout();
        Text = AppTitle;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { /* 既定のアイコンのまま */ }
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(1080, 640);
        MinimumSize = new Size(760, 420);
        KeyPreview = true;

        BuildLayout();
        WireEvents();
        ResumeLayout(false);
        PerformLayout();

        _optPageNumbers.Checked = startup.IsBatch ? startup.PageNumbers : _settings.PageNumbers;
        _optToc.Checked = startup.IsBatch ? startup.TableOfContents : _settings.TableOfContents;
        _optDuplex.Checked = startup.IsBatch ? startup.DuplexBlanks : _settings.DuplexBlanks;
        _optOpen.Checked = startup.IsBatch ? startup.Open : _settings.OpenAfterCreate;

        RefreshAll();
        SetStatus("ファイルを追加すると、すぐ裏で PDF に変換します。並べ替えて「結合して PDF を作成」を押してください（結合しないときは「個別に PDF を保存」）。");
    }

    /// <summary>スクリプトから使ったとき（--out）の終了コード。</summary>
    public int ExitCode { get; private set; }

    private static Button ToolButton(string text) => new()
    {
        Text = text,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        Padding = new Padding(6, 2, 6, 2),
        Margin = new Padding(0, 0, 4, 4),
    };

    private static CheckBox Option(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(0, 3, 18, 3) };

    private static Control Gap() => new Panel { Width = 10, Height = 1, Margin = Padding.Empty };

    private void BuildLayout()
    {
        // 一覧
        _icons.ImageSize = SystemInformation.SmallIconSize;
        _list.SmallImageList = _icons;
        _list.Columns.AddRange([_colName, _colTitle, _colPages, _colRange, _colState, _colOrder]);
        _colOrder.DisplayIndex = 0;
        _list.Controls.Add(_emptyHint);
        _list.Controls.Add(_insertMark);
        _list.ContextMenuStrip = BuildContextMenu();

        _split.Panel1.Controls.Add(_list);
        _split.Panel2.Controls.Add(_preview);
        _preview.BorderStyle = BorderStyle.FixedSingle;

        // 上のボタン
        // 幅が足りないとき（表示倍率が大きい画面や、狭いウィンドウ）は、ボタンを切らずに 2 段に折り返す
        var toolbar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = new Padding(0, 0, 0, 4) };
        toolbar.Controls.AddRange([_addFiles, _addFolder, Gap(), _moveUp, _moveDown, _sortByName, Gap(), _remove, _clear, Gap(), _undo]);

        // オプション
        var options = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 0) };
        options.Controls.AddRange([_optPageNumbers, _optToc, _optDuplex, _optOpen]);

        // 下の行: ふだんは合計、作成中は進み具合
        // 右の大きいボタンと高さがそろうよう、行の中で上下中央に置く
        var gauge = new Panel { Anchor = AnchorStyles.Left | AnchorStyles.Right, Height = 24, Margin = new Padding(0, 0, 10, 0) };
        gauge.Controls.Add(_summary);
        gauge.Controls.Add(_progress);

        var bottom = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 4, Margin = new Padding(0, 4, 0, 0) };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.Controls.Add(gauge, 0, 0);
        bottom.Controls.Add(_cancel, 1, 0);
        bottom.Controls.Add(_runEach, 2, 0);
        bottom.Controls.Add(_run, 3, 0);

        // 最後に押すボタンなので、ほかより目立たせる（個別の保存は、それに並ぶ控えめな選択肢）
        _run.Font = new Font(Font, FontStyle.Bold);
        _run.Padding = new Padding(20, 6, 20, 6);
        _runEach.Padding = new Padding(10, 6, 10, 6);
        _cancel.Padding = new Padding(10, 6, 10, 6);

        _status.Margin = new Padding(0, 6, 0, 0);
        _status.LinkColor = _status.VisitedLinkColor = Theme.Link;
        _status.ActiveLinkColor = Theme.Link;

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(toolbar);
        root.Controls.Add(_split);
        root.Controls.Add(options);
        root.Controls.Add(_status);
        root.Controls.Add(bottom);
        Controls.Add(root);

        _toolTip.SetToolTip(_runEach, "結合せずに、ファイルごとに「元のファイル名.pdf」を保存（元のファイルと同じフォルダか、選んだフォルダ）");
        _toolTip.SetToolTip(_addFiles, "結合するファイルを選ぶ（Ctrl+O）");
        _toolTip.SetToolTip(_addFolder, "フォルダの中の対象ファイルを、サブフォルダも含めてすべて追加");
        _toolTip.SetToolTip(_moveUp, "選んだ行を 1 つ上へ（Alt+↑）。行のドラッグでも動かせます");
        _toolTip.SetToolTip(_moveDown, "選んだ行を 1 つ下へ（Alt+↓）");
        _toolTip.SetToolTip(_sortByName, "ファイル名の順（エクスプローラーと同じ数字の並び）に並べ替え");
        _toolTip.SetToolTip(_remove, "選んだ行をリストから外す（Delete）。ファイルは消えません");
        _toolTip.SetToolTip(_undo, "リストの変更を 1 つ戻す（Ctrl+Z）。やり直しは Ctrl+Y");
        _toolTip.SetToolTip(_optPageNumbers, "各ページの下に「3 / 25」の形で通し番号を入れます");
        _toolTip.SetToolTip(_optToc, "先頭に、しおり名とページ番号の一覧を付けます。行を押すとそのページへ飛べます");
        _toolTip.SetToolTip(_optDuplex, "各ファイルが紙の表（奇数ページ）から始まるよう、必要なところに白紙を入れます");
        _toolTip.SetToolTip(_optOpen, "できあがった PDF をすぐに開きます");
        _toolTip.SetToolTip(_run, "上から順に 1 つの PDF にまとめます（Ctrl+Enter）");
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("開く", null, (_, _) => OpenSelected());
        menu.Items.Add("フォルダで表示", null, (_, _) => RevealSelected());
        var openPdf = menu.Items.Add("変換結果の PDF を開く", null, (_, _) => OpenConvertedSelected());
        menu.Items.Add(new ToolStripSeparator());
        var rename = new ToolStripMenuItem("しおり名を変更", null, (_, _) => EditSelected(ColTitle)) { ShortcutKeyDisplayString = "F2" };
        var range = new ToolStripMenuItem("ページを指定...", null, (_, _) => EditSelected(ColRange)) { ShortcutKeyDisplayString = "F3" };
        var sheets = new ToolStripMenuItem("シートと幅の設定...", null, (_, _) => ChooseSheetsForSelected());
        var password = new ToolStripMenuItem("パスワードを入力...", null, (_, _) => AskPasswordForSelected());
        var retry = new ToolStripMenuItem("もう一度変換", null, (_, _) => RetrySelected());
        var extract = new ToolStripMenuItem("ページを選んで保存...", null, (_, _) => ExtractSelected());
        var split = new ToolStripMenuItem("分割...", null, (_, _) => SplitSelected());
        menu.Items.AddRange([rename, range, sheets, password, retry, extract, split]);
        menu.Items.Add(new ToolStripSeparator());
        var remove = new ToolStripMenuItem("削除", null, (_, _) => RemoveSelected()) { ShortcutKeyDisplayString = "Delete" };
        menu.Items.Add(remove);

        menu.Opening += (_, e) =>
        {
            var selected = SelectedEntries().ToList();
            if (selected.Count == 0)
            {
                e.Cancel = true;
                return;
            }
            var single = selected.Count == 1 ? selected[0] : null;
            bool idle = !_busy;
            openPdf.Enabled = single?.Converted is not null;
            rename.Enabled = idle && single is not null;
            range.Enabled = idle && single is not null;
            sheets.Visible = single?.Kind == FileKind.Excel;
            sheets.Enabled = idle && single?.Converted?.Sheets is not null;
            password.Visible = single?.State == EntryState.NeedsPassword;
            password.Enabled = idle;
            retry.Enabled = idle && selected.Any(s => !s.IsBusy);
            split.Enabled = idle && single is { IsUsable: true };
            extract.Enabled = idle && single?.Converted is not null;
            remove.Enabled = idle;
        };
        return menu;
    }

    private void WireEvents()
    {
        _addFiles.Click += (_, _) => ChooseFiles();
        _addFolder.Click += (_, _) => ChooseFolder();
        _moveUp.Click += (_, _) => MoveSelectedBy(-1);
        _moveDown.Click += (_, _) => MoveSelectedBy(+1);
        _sortByName.Click += (_, _) => SortByName();
        _remove.Click += (_, _) => RemoveSelected();
        _clear.Click += (_, _) => ClearAll();
        _undo.Click += (_, _) => Undo();
        _run.Click += async (_, _) => await CreateAsync(null);
        _runEach.Click += async (_, _) => await SaveEachAsync();
        _cancel.Click += (_, _) => CancelCreate();
        _emptyHint.Click += (_, _) => ChooseFiles();
        _status.LinkClicked += OnStatusLinkClicked;

        foreach (var option in new[] { _optPageNumbers, _optToc, _optDuplex })
            option.CheckedChanged += (_, _) => UpdateSummary();

        _list.SelectedIndexChanged += (_, _) => OnSelectionChanged();
        _list.KeyDown += OnListKeyDown;
        _list.MouseDoubleClick += OnListDoubleClick;
        _list.ItemDrag += OnItemDrag;
        _list.Resize += (_, _) => FitColumns();
        DpiChangedAfterParent += (_, _) => FitColumns();

        foreach (Control target in new Control[] { _list, _emptyHint })
        {
            target.DragEnter += OnDragOver;
            target.DragOver += OnDragOver;
            target.DragLeave += (_, _) => HideInsertMark();
            target.DragDrop += OnDragDrop;
        }

        _preview.PageToggled += TogglePage;
        _preview.PageZoomRequested += ShowPage;
        _preview.PageRotateRequested += RotatePage;
        _preview.ChooseSheetsRequested += ChooseSheets;
        _preview.SplitRequested += Split;
        _preview.ExtractRequested += ExtractPages;
        _preview.PasswordRequested += AskPassword;
        _preview.RetryRequested += entry => Retry([entry]);
        _preview.UseAllPagesRequested += entry => SetPageRange(entry, "");
        _preview.ResetSheetsRequested += entry =>
        {
            PushUndo();
            entry.Sheets = null;
            Prepare(entry);
        };
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        // 文字の高さは表示倍率どおりなので、倍率の調整が済んだここで決める
        _status.Height = _status.PreferredHeight + LogicalToDeviceUnits(6);
        // 保存してある位置は実際の画面のピクセルなので、表示倍率の調整が済んだここで当てる
        RestoreWindow();
        // 既定の大きさが画面より大きいとき（小さい画面で表示倍率が大きいなど）は、画面に収める
        var area = Screen.FromControl(this).WorkingArea;
        if (WindowState == FormWindowState.Normal && (Width > area.Width || Height > area.Height))
            Bounds = new Rectangle(area.Left, area.Top, Math.Min(Width, area.Width), Math.Min(Height, area.Height));
        try
        {
            int width = _split.ClientSize.Width;
            _split.Panel1MinSize = LogicalToDeviceUnits(320);
            _split.Panel2MinSize = LogicalToDeviceUnits(200);
            int distance = (int)(width * Math.Clamp(_settings.SplitRatio, 0.3, 0.85));
            _split.SplitterDistance = Math.Clamp(distance, _split.Panel1MinSize, Math.Max(_split.Panel1MinSize, width - _split.Panel2MinSize - _split.SplitterWidth));
        }
        catch { /* 幅が足りないときは既定の位置のまま */ }
        FitColumns();
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        string[] early;
        lock (_earlyPaths)
        {
            _shown = true;
            early = [.. _earlyPaths];
            _earlyPaths.Clear();
        }
        if (_startup.Paths.Length + early.Length > 0)
            await AddPathsAsync([.. _startup.Paths, .. early], null);

        if (_startup.Output is { } output)
        {
            bool ok = _list.Items.Count > 0 && await CreateAsync(output);
            if (_list.Items.Count == 0) Console.Error.WriteLine("結合できるファイルが指定されていません。");
            ExitCode = ok ? 0 : 1;
            Close();
        }
    }

    protected override void OnActivated(EventArgs e)
    {
        base.OnActivated(e);
        // ほかのアプリでファイルを直して戻ってきたら、変わったものだけ変換し直す
        if (IsHandleCreated && !_busy) _ = CheckSourcesAsync();
    }

    /// <summary>
    /// 列の幅を、一覧の幅にちょうど収まるよう配る（横のスクロールを出さない）。
    /// 幅が足りないときは、ページ数とページ指定の列も縮める。
    /// </summary>
    private void FitColumns()
    {
        int width = _list.ClientSize.Width - 1;
        if (width <= 0) return;
        // 縦のスクロールバーがまだなくても、行が増えて出たときに横へはみ出さないよう、その分を空けておく
        if (!ListHasVerticalScrollBar()) width -= SystemInformation.VerticalScrollBarWidth;

        int order = LogicalToDeviceUnits(_list.Items.Count >= 100 ? 40 : 32);
        int pages = LogicalToDeviceUnits(58), range = LogicalToDeviceUnits(84);
        int fixedWidth = order + pages + range;
        if (width < fixedWidth * 2)
        {
            pages = width * 58 / 142 / 2;
            range = width * 84 / 142 / 2;
            fixedWidth = order + pages + range;
        }
        int flexible = Math.Max(0, width - fixedWidth);
        _list.BeginUpdate();
        _colOrder.Width = order;
        _colPages.Width = pages;
        _colRange.Width = range;
        // 一番大事なファイル名に多く配る。状態は「OK」で済むことが多く、長い理由は行のツールチップで全文が見られる
        _colName.Width = flexible * 44 / 100;
        _colTitle.Width = flexible * 32 / 100;
        _colState.Width = flexible - _colName.Width - _colTitle.Width;
        _list.EndUpdate();
    }

    private bool ListHasVerticalScrollBar()
    {
        const int GwlStyle = -16, WsVScroll = 0x00200000;
        return _list.IsHandleCreated && (NativeMethods.GetWindowLong(_list.Handle, GwlStyle) & WsVScroll) != 0;
    }

    private void SetStatus(string text)
    {
        _status.Links.Clear();
        _status.Text = text;
        _status.LinkArea = new LinkArea(0, 0);
        _toolTip.SetToolTip(_status, text);
    }

    /// <summary>作成できたことを、開くためのリンク付きで出す。</summary>
    private void SetCreatedStatus(string output, int files, int pages)
    {
        _lastOutput = output;
        const string open = "開く", reveal = "フォルダで表示";
        string head = $"作成しました: {Path.GetFileName(output)}（{files} ファイル、{pages} ページ）　";
        _status.Links.Clear();
        _status.Text = $"{head}{open}　{reveal}";
        _status.Links.Add(head.Length, open.Length, "open");
        _status.Links.Add(head.Length + open.Length + 1, reveal.Length, "reveal");
        _toolTip.SetToolTip(_status, output);
    }

    /// <summary>複数のファイルを書き出したこと（分割、個別の保存）を、保存先を開くためのリンク付きで出す。</summary>
    private void SetFilesStatus(IReadOnlyList<string> files, string done)
    {
        _lastOutput = files[0];
        // 保存先の長いパスを行に入れるとリンクが見えなくなるので、パスはツールチップに回す
        const string reveal = "フォルダで表示";
        string head = $"{done}（{Path.GetFileName(files[0])} など）　";
        _status.Links.Clear();
        _status.Text = head + reveal;
        _status.Links.Add(head.Length, reveal.Length, "reveal");
        _toolTip.SetToolTip(_status, Path.GetDirectoryName(files[0]) + "\r\n" + string.Join("\r\n", files.Take(20).Select(Path.GetFileName)));
    }

    private void OnStatusLinkClicked(object? sender, LinkLabelLinkClickedEventArgs e)
    {
        if (_lastOutput is null) return;
        if (e.Link?.LinkData as string == "open") ShellOpen(_lastOutput);
        else Reveal(_lastOutput);
    }

    private MergeOptions CurrentOptions(string documentTitle) =>
        new(_optPageNumbers.Checked, _optToc.Checked, _optDuplex.Checked, documentTitle);

    private void UpdateSummary()
    {
        var entries = Entries().ToList();
        if (entries.Count == 0)
        {
            _summary.Text = "";
            return;
        }

        var usable = entries.Where(e => e.IsUsable).ToList();
        int busy = entries.Count(e => e.IsBusy);
        int failed = entries.Count - usable.Count - busy;
        var options = CurrentOptions("");
        var (total, _, blankBefore) = PdfMerger.Plan(usable.Select(e => e.SelectedPages.Count).ToList(), options);
        int toc = options.TableOfContents && usable.Count > 0 ? PdfMerger.TocPageCount(usable.Count) : 0;
        int blanks = blankBefore.Count(b => b);

        var parts = new List<string> { $"{entries.Count} ファイル" };
        if (usable.Count > 0)
        {
            // 足されるページがあるときだけ、その内訳を添える
            var added = new List<string>();
            if (toc > 0) added.Add($"目次 {toc}");
            if (blanks > 0) added.Add($"白紙 {blanks}");
            parts.Add(added.Count == 0 ? $"{total} ページ" : $"仕上がり {total} ページ（{string.Join("・", added)} を含む）");
        }
        if (busy > 0) parts.Add($"変換中 {busy} 件");
        if (failed > 0) parts.Add($"結合できない行 {failed} 件");
        _summary.Text = string.Join("　・　", parts);
        _summary.ForeColor = failed > 0 ? Theme.Danger : SystemColors.ControlText;
    }

    private void UpdateButtons()
    {
        bool idle = !_busy;
        int count = _list.Items.Count;
        var indices = _list.SelectedIndices.Cast<int>().ToList();
        int selected = indices.Count;

        _addFiles.Enabled = idle;
        _addFolder.Enabled = idle;
        _moveUp.Enabled = idle && selected > 0 && indices.Min() > 0;
        _moveDown.Enabled = idle && selected > 0 && indices.Max() < count - 1;
        _sortByName.Enabled = idle && count > 1;
        _remove.Enabled = idle && selected > 0;
        _clear.Enabled = idle && count > 0;
        _undo.Enabled = idle && _undoStack.Count > 0;
        // フォルダの中を探している間は、見つかったものがまだ一覧に入っていないので作成させない
        _run.Enabled = idle && count > 0 && !IsScanning;
        _runEach.Enabled = _run.Enabled;
        _cancel.Enabled = (_busy && _cts is { IsCancellationRequested: false }) || IsScanning;
        _emptyHint.Cursor = idle ? Cursors.Hand : Cursors.Default;
        foreach (var option in new[] { _optPageNumbers, _optToc, _optDuplex, _optOpen })
            option.Enabled = idle;
    }

    private void ShellOpen(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })?.Dispose(); }
        catch (Exception ex) { ShowError("開けませんでした", ex.Message); }
    }

    private void Reveal(string path)
    {
        try { Process.Start("explorer.exe", $"/select,\"{path}\"")?.Dispose(); }
        catch (Exception ex) { ShowError("フォルダを開けませんでした", ex.Message); }
    }

    private void ShowError(string heading, string text)
    {
        if (_startup.IsBatch)
        {
            Console.Error.WriteLine($"{heading}: {text}");
            return;
        }
        TaskDialog.ShowDialog(this, new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = heading,
            Text = text,
            Icon = TaskDialogIcon.Error,
        });
    }

    /// <summary>あとから起動された自分が送ってきたファイルを、このウィンドウのリストに加える。</summary>
    public void ReceivePaths(string[] paths)
    {
        if (IsDisposed) return;
        lock (_earlyPaths)
        {
            if (!_shown)
            {
                _earlyPaths.AddRange(paths);
                return;
            }
        }
        try
        {
            BeginInvoke(() =>
            {
                if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
                Activate();
                if (paths.Length == 0) return;
                if (_busy) SetStatus("作成中のため、ファイルを追加できませんでした。終わってからもう一度追加してください。");
                else _ = AddPathsAsync(paths, null);
            });
        }
        catch (InvalidOperationException) { /* 閉じている途中 */ }
    }

    private void RestoreWindow()
    {
        if (_settings.WindowWidth <= 0 || _settings.WindowHeight <= 0) return;
        var bounds = new Rectangle(_settings.WindowX, _settings.WindowY, _settings.WindowWidth, _settings.WindowHeight);
        if (!Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(bounds))) return;
        StartPosition = FormStartPosition.Manual;
        Bounds = bounds;
        if (_settings.Maximized) WindowState = FormWindowState.Maximized;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _scan?.Cancel(); // フォルダの中を探している途中なら止める
        if (_busy && !_startup.IsBatch)
        {
            e.Cancel = true;
            if (_closeWhenStopped) return;
            var answer = MessageBox.Show(this, "PDF を作成中です。中止して閉じますか？", AppTitle,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (answer == DialogResult.Yes)
            {
                _closeWhenStopped = true;
                CancelCreate();
            }
            return;
        }

        if (!_startup.IsBatch)
        {
            var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            _settings.WindowX = bounds.X;
            _settings.WindowY = bounds.Y;
            _settings.WindowWidth = bounds.Width;
            _settings.WindowHeight = bounds.Height;
            _settings.Maximized = WindowState == FormWindowState.Maximized;
            if (_split.ClientSize.Width > 0)
                _settings.SplitRatio = (double)_split.SplitterDistance / _split.ClientSize.Width;
            _settings.PageNumbers = _optPageNumbers.Checked;
            _settings.TableOfContents = _optToc.Checked;
            _settings.DuplexBlanks = _optDuplex.Checked;
            _settings.OpenAfterCreate = _optOpen.Checked;
            _settings.Save();
        }
        base.OnFormClosing(e);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        base.OnFormClosed(e);
        // 変換中の Office を止めて片付ける間、ウィンドウは先に消しておく
        Hide();
        foreach (var entry in Entries()) entry.Pending?.Cancel();
        _service.Dispose();
        _preview.Dispose();
        _workspace.Dispose();
    }
}

/// <summary>行の更新でちらつかないようにダブルバッファを有効にした ListView。</summary>
internal sealed class BufferedListView : ListView
{
    public BufferedListView() => DoubleBuffered = true;
}
