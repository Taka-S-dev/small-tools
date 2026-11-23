using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>しおりの選び方。ダイアログの中で書き換え、OK のときだけ一覧と設定に移す。</summary>
internal sealed class BookmarkChoices
{
    /// <summary>元のしおりを何段まで付けるか（全体の既定）。null ならすべて、0 なら付けない。</summary>
    public int? Depth { get; set; }

    /// <summary>フォルダを入れたときに、フォルダごとにまとめるか。</summary>
    public bool Folders { get; set; }

    /// <summary>ファイルの番号 → しおり名。空なら元のファイル名に戻す。</summary>
    public Dictionary<int, string> Titles { get; } = [];

    /// <summary>ファイルの番号 → そのファイルだけの段数（MergePart.OutlineDepth と同じ意味。null なら全体に合わせる）。</summary>
    public Dictionary<int, int?> FileDepths { get; } = [];

    /// <summary>フォルダの Key → 付け直した名前。</summary>
    public Dictionary<string, string> FolderNames { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>ファイルの番号 → 元のしおりの手直し（MergePart.OutlineEdits と同じ意味）。</summary>
    public Dictionary<int, Dictionary<string, string?>> OutlineEdits { get; } = [];

    /// <summary>ファイルの番号 → ページを見て足したしおり（MergePart.Added と同じ意味）。</summary>
    public Dictionary<int, List<AddedBookmark>> Added { get; } = [];

    /// <summary>そのファイルの手直しを書き込む先（なければ作る）。</summary>
    public Dictionary<string, string?> EditsOf(int part) =>
        OutlineEdits.TryGetValue(part, out var edits) ? edits : OutlineEdits[part] = [];

    /// <summary>ファイルの番号 → 元のしおりの段の付け直し（MergePart.OutlineLevels と同じ意味）。</summary>
    public Dictionary<int, Dictionary<string, int>> OutlineLevels { get; } = [];

    /// <summary>そのファイルの段の付け直しを書き込む先（なければ作る）。</summary>
    public Dictionary<string, int> LevelsOf(int part) =>
        OutlineLevels.TryGetValue(part, out var levels) ? levels : OutlineLevels[part] = [];
}

/// <summary>しおりの確認で右に並べる、1 ファイルのページ。Pages は結合に使うページ（元のファイルでの番号）。</summary>
internal sealed record BookmarkPages(ConvertedPdf Pdf, IReadOnlyList<int> Pages, IReadOnlyDictionary<int, int> Rotations);

/// <summary>
/// 結合した PDF に付くしおりを、作る前に木の形で見せる。右には、木で選んだファイルのページが並び、
/// ページを右クリックするとそのページにしおりを足せる（足したしおりはすぐ木に出て、そのまま名前を書ける）。
/// ファイル・フォルダ・しおりの名前はその場で変えられ、しおりは段を上げ下げしたり外したりできる（Ctrl / Shift でまとめて選べる）。
/// 元のしおりを何段まで付けるかは、全体とファイルごとに選べる。変えるたびに木を作り直す（結合と同じ決まりで求める）。
/// </summary>
internal sealed class BookmarkDialog : Form
{
    private static readonly (string Text, int? Depth)[] Depths =
        [("すべて", null), ("1 段まで", 1), ("2 段まで", 2), ("3 段まで", 3), ("付けない", 0)];

    // ファイルごとの段数。null は「全体の設定に合わせる」
    private static readonly (string Text, int? Depth)[] FileDepthChoices =
        [("全体の設定に合わせる", null), ("すべて", MergePart.AllLevels), ("1 段まで", 1), ("2 段まで", 2), ("3 段まで", 3), ("付けない", 0)];

    private readonly Func<BookmarkChoices, IReadOnlyList<Bookmark>> _build;
    private readonly BookmarkChoices _choices;
    private readonly Func<int, BookmarkPages?> _pagesOf;
    private readonly Func<BookmarkChoices, ConvertedPdf?>? _contents;
    /// <summary>目次ページを並べているときの _shownPart。</summary>
    private const int ContentsPart = -2;
    private readonly Action<int, int, IWin32Window>? _zoom;
    private readonly PageStrip _strip = new() { Dock = DockStyle.Fill };
    private readonly Label _pagesTitle = new() { AutoSize = true, AutoEllipsis = false, Margin = new Padding(0, 0, 0, 4) };
    private readonly Label _pagesNote = new() { AutoSize = true, Margin = new Padding(0, 4, 0, 0) };
    private readonly SplitContainer _split = new() { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
    /// <summary>右に並べているファイルの番号。まだなければ -1。</summary>
    private int _shownPart = -1;
    /// <summary>右のページ（元のファイルでの番号）→ 結合した PDF でのページ番号（1 始まり）。結合しないページは入っていない。</summary>
    private Dictionary<int, int> _mergedNumbers = [];
    /// <summary>木を作り直している間や、右のページを押して行を選んだときは、選び直しで右のページを動かさない。</summary>
    private bool _rebuilding;
    private readonly ComboBox _depth = new() { DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left, Width = 120 };
    private readonly CheckBox _folders = new()
    {
        AutoSize = true,
        Text = "フォルダを入れたときは、フォルダごとにまとめる",
        Margin = new Padding(0, 6, 0, 6),
    };
    // 名前は TreeView の LabelEdit ではなく、行の上に重ねる入力欄で変える（一覧の「しおり名」と同じ）。
    // LabelEdit は、画面の読み上げなどが動いていると、編集を終えた後にアプリが落ちることがある（.NET 10.0.12 で確認）
    private readonly TreeView _tree = new() { Dock = DockStyle.Fill, HideSelection = false, ShowNodeToolTips = true };
    private TextBox? _editor;
    // 選んだしおりへの操作のボタン（右クリックやキーを知らなくても使えるように）
    private readonly Button _promote = ExcelSheetsDialog.DialogButton("← 段を上げる");
    private readonly Button _demote = ExcelSheetsDialog.DialogButton("段を下げる →");
    private readonly Button _rename = ExcelSheetsDialog.DialogButton("名前を変える");
    private readonly Button _remove = ExcelSheetsDialog.DialogButton("外す");
    private readonly ToolTip _toolTip = new();
    /// <summary>
    /// Ctrl / Shift を押しながらクリックして、選んでいる行と一緒に選んだ行。
    /// 木は作り直すたびに行が変わるので、行の作り（KeyOf）で覚える。
    /// </summary>
    private readonly HashSet<string> _marked = [];
    /// <summary>次の選び直しで、一緒に選んだ行を消さない（Ctrl / Shift のクリック）。</summary>
    private bool _keepMarks;
    /// <summary>次の選び直しを取りやめる（Ctrl で、選んでいない行の印を外すとき）。</summary>
    private bool _cancelSelect;
    private readonly ContextMenuStrip _menu = new();
    private TreeNode? _menuTarget;
    private bool _menuByMouse;
    // フォルダの段にだけフォルダのアイコンを付ける（ほかの行は、幅をそろえるための空の絵）
    private readonly ImageList _icons = new() { ColorDepth = ColorDepth.Depth32Bit };
    private const int NoIcon = 0, FolderIcon = 1;

    /// <param name="build">選び方 → できあがるしおり。</param>
    /// <param name="choices">今の選び方。ダイアログの中で書き換えるので、呼ぶ側は写しを渡す。</param>
    /// <param name="note">一覧のうち、しおりに入らないファイルがあるときの説明（変換中や、結合できないもの）。なければ空。</param>
    /// <param name="pagesOf">ファイルの番号 → 右に並べるページ。並べられなければ null。</param>
    /// <param name="zoom">(ファイルの番号, ページ, 親) → ページを大きく見せる。null なら見せない。</param>
    /// <param name="contents">選び方 → 結合するときに作る目次ページの見本。目次を付けないときは null。</param>
    public BookmarkDialog(Func<BookmarkChoices, IReadOnlyList<Bookmark>> build, BookmarkChoices choices, string note,
        Func<int, BookmarkPages?> pagesOf, Action<int, int, IWin32Window>? zoom = null, Func<BookmarkChoices, ConvertedPdf?>? contents = null)
    {
        _contents = contents;
        _build = build;
        _choices = choices;
        _pagesOf = pagesOf;
        _zoom = zoom;
        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = "しおりの確認";
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        ShowInTaskbar = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(980, 640);
        MinimumSize = new Size(640, 400);

        _tree.ContextMenuStrip = _menu;

        foreach (var (text, _) in Depths) _depth.Items.Add(text);
        int index = Array.FindIndex(Depths, d => d.Depth == choices.Depth);
        _depth.SelectedIndex = index >= 0 ? index : 0;
        _folders.Checked = choices.Folders;

        var depthRow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = Padding.Empty };
        depthRow.Controls.Add(new Label
        {
            Text = "元のファイルのしおり（Word の見出しなど）",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 6, 8, 0),
        });
        depthRow.Controls.Add(_depth);

        var ok = ExcelSheetsDialog.DialogButton("OK", DialogResult.OK);
        var cancel = ExcelSheetsDialog.DialogButton("キャンセル", DialogResult.Cancel);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 10, 0, 0) };
        buttons.Controls.AddRange([cancel, ok]);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(12) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(depthRow);
        root.Controls.Add(_folders);
        // 列の幅を指定しないと中身に合わせて広がり、右のページが領域からはみ出すので、領域いっぱいに決める
        var treePane = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Margin = Padding.Empty, Padding = new Padding(0, 0, 4, 0) };
        treePane.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        treePane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        treePane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        treePane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        treePane.Controls.Add(new Label { Text = "できあがるしおり（括弧の中は、結合した PDF でのページ）", AutoSize = true, Margin = new Padding(0, 6, 0, 4) });
        var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = new Padding(0, 0, 0, 4) };
        foreach (var button in new[] { _promote, _demote, _rename, _remove }) button.Margin = new Padding(0, 0, 6, 0);
        actions.Controls.AddRange([_promote, _demote, _rename, _remove]);
        treePane.Controls.Add(actions);
        treePane.Controls.Add(_tree);
        _split.Panel1.Controls.Add(treePane);

        var pagesPane = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Margin = Padding.Empty, Padding = new Padding(4, 0, 0, 0) };
        pagesPane.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        pagesPane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        pagesPane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        pagesPane.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _pagesTitle.Margin = new Padding(0, 6, 0, 4);
        pagesPane.Controls.Add(_pagesTitle);
        pagesPane.Controls.Add(_strip);
        pagesPane.Controls.Add(_pagesNote);
        _split.Panel2.Controls.Add(pagesPane);
        root.Controls.Add(_split);

        root.Controls.Add(new Label
        {
            Text = "上のボタンで、選んだしおりの段の上げ下げ・名前の変更・外すができます（Alt+Shift+← →、F2、Delete でも）。"
                + "右のページを押すとそのページのしおりを選び、右クリックするとそのページにしおりを付けられます。"
                + "木の右クリックで、ファイルごとに元のしおりを何段まで付けるかも選べます。",
            AutoSize = true,
            MaximumSize = new Size(940, 0),
            ForeColor = SystemColors.GrayText,
            Margin = new Padding(0, 6, 0, 0),
        });
        root.Controls.Add(new Label
        {
            Text = note,
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Margin = new Padding(0, 6, 0, 0),
            Visible = note.Length > 0,
        });
        root.Controls.Add(buttons);
        Controls.Add(root);

        // 画面の読み上げや自動操作で見分けられるよう、名前を付ける
        (_depth.Name, _depth.AccessibleName) = ("Depth", "元のファイルのしおり");
        (_tree.Name, _tree.AccessibleName) = ("Bookmarks", "できあがるしおり");

        AcceptButton = ok;
        CancelButton = cancel;
        _depth.SelectedIndexChanged += (_, _) =>
        {
            _choices.Depth = Depths[Math.Max(0, _depth.SelectedIndex)].Depth;
            Rebuild();
        };
        _folders.CheckedChanged += (_, _) =>
        {
            _choices.Folders = _folders.Checked;
            Rebuild();
        };
        _promote.Click += (_, _) => MoveLevel(Chosen(), -1);
        _demote.Click += (_, _) => MoveLevel(Chosen(), +1);
        _rename.Click += (_, _) => { if (_tree.SelectedNode is { } node) BeginInvoke(() => BeginRename(node)); };
        _remove.Click += (_, _) => Drop(Chosen());
        const string Several = "Ctrl か Shift を押しながらクリックすると、いくつかまとめて選べます。";
        _toolTip.SetToolTip(_promote, "選んだしおりを 1 段上げます。下のしおりも一緒に動きます（Alt+Shift+←）。" + Several);
        _toolTip.SetToolTip(_demote, "選んだしおりを、すぐ上のしおりの下へ 1 段下げます。下のしおりも一緒に動きます（Alt+Shift+→）。" + Several);
        _toolTip.SetToolTip(_rename, "選んだしおりの名前を変えます（F2）");
        _toolTip.SetToolTip(_remove, "選んだしおりを外します。元のしおりは下のしおりも外れます（Delete）。" + Several);
        _tree.KeyDown += (_, e) =>
        {
            _menuByMouse = false; // この後の Shift+F10 は、選んでいる行のメニューにする
            if (e.Alt && e.Shift && e.KeyCode is Keys.Left or Keys.Right)
            {
                e.Handled = e.SuppressKeyPress = true; // 木の開け閉めはしない
                MoveLevel(Chosen(), e.KeyCode == Keys.Left ? -1 : +1);
            }
            else if (e.KeyCode == Keys.F2 && _tree.SelectedNode is { } node) BeginRename(node);
            else if (e.KeyCode == Keys.Delete) Drop(Chosen());
        };
        // 右クリックした行を選んでから、メニューを開く（TreeView は右クリックでは選ばない）。
        // 行のないところで右クリックしたときは、前に選んでいた行のメニューを出さない
        _tree.MouseDown += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                // 開け閉めの印を押しただけなら、選び方は変えない
                if (_tree.HitTest(e.Location).Location != TreeViewHitTestLocations.PlusMinus) OnLeftDown(_tree.GetNodeAt(e.Location));
                return;
            }
            if (e.Button != MouseButtons.Right) return;
            _menuTarget = _tree.GetNodeAt(e.Location);
            _menuByMouse = true;
            // 一緒に選んだ行の上での右クリックなら、まとめて選んだまま操作する
            if (_menuTarget is not null && !IsChosen(_menuTarget)) _marked.Clear();
            else if (_menuTarget is not null) _keepMarks = true;
            if (_menuTarget is not null) _tree.SelectedNode = _menuTarget;
            ApplyMarks();
        };
        _tree.BeforeSelect += (_, e) =>
        {
            e.Cancel = _cancelSelect;
            _cancelSelect = false;
        };
        _menu.Opening += OnMenuOpening;
        _tree.AfterSelect += (_, e) =>
        {
            // 何も押さずに選び直したら（クリックや矢印キー）、一緒に選んだ行は外す
            if (!_keepMarks && !_rebuilding) _marked.Clear();
            _keepMarks = false;
            ApplyMarks();
            UpdateActions();
            if (!_rebuilding && e.Node is not null) ShowPagesOf(e.Node, scroll: true);
        };
        _strip.PageClickSelects = true;
        _strip.PageClicked += SelectBookmarkAt;
        _strip.PageMark = page => string.Join("・", (_choices.Added.GetValueOrDefault(_shownPart) ?? []).Where(b => b.Page == page).Select(b => b.Title));
        _strip.PageBookmarkRequested += AddAt;
        _strip.PageZoomRequested += page =>
        {
            if (_shownPart >= 0) _zoom?.Invoke(_shownPart, page, this);
        };
        ResumeLayout(false);
        PerformLayout();
        Rebuild();
        // 最初のファイルのページを並べておく
        if (FirstFile(_tree.Nodes) is { } first)
        {
            _tree.SelectedNode = first;
            ShowPagesOf(first, scroll: false);
        }
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        // 表示倍率の調整が済んだ幅で、左（しおり）と右（ページ）を分ける
        _split.SplitterDistance = _split.Width * 45 / 100;
        _tree.Indent = LogicalToDeviceUnits(16);
        LoadIcons();
        FitRows();
    }

    // 表示倍率の違う画面へ動かしたら、アイコンと行の高さを合わせ直す
    protected override void OnDpiChanged(DpiChangedEventArgs e)
    {
        base.OnDpiChanged(e);
        _tree.Indent = LogicalToDeviceUnits(16);
        LoadIcons();
        FitRows();
    }

    /// <summary>
    /// フォルダのアイコンを、このダイアログがいる画面の表示倍率の大きさで用意する。
    /// SystemInformation.SmallIconSize は Windows 全体の倍率で決まり、倍率の違う画面では大きすぎて行が間延びする。
    /// </summary>
    private void LoadIcons()
    {
        var size = new Size(LogicalToDeviceUnits(16), LogicalToDeviceUnits(16));
        bool attached = _tree.ImageList is not null;
        _tree.ImageList = null;
        _icons.Images.Clear();
        _icons.ImageSize = size;
        _icons.Images.Add(new Bitmap(size.Width, size.Height)); // NoIcon: 幅をそろえるための空の絵
        using (var folder = ShellIcons.ForFolder()) _icons.Images.Add(folder?.ToBitmap() ?? new Bitmap(size.Width, size.Height));
        if (attached) _tree.ImageList = _icons;
    }

    /// <summary>
    /// 行の高さを、文字（とフォルダのアイコン）が収まる高さにする。指定しないと既定の高さが大きく、間延びして見える。
    /// アイコンの一覧を付け外しすると高さが戻るので、そのたびに呼ぶ。
    /// </summary>
    private void FitRows()
    {
        if (!IsHandleCreated) return;
        int height = _tree.Font.Height + LogicalToDeviceUnits(6);
        if (_tree.ImageList is not null) height = Math.Max(height, _icons.ImageSize.Height + LogicalToDeviceUnits(4));
        _tree.ItemHeight = height;
    }

    private static TreeNode? FirstFile(TreeNodeCollection nodes)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Tag is Bookmark { Kind: BookmarkKind.File }) return node;
            if (FirstFile(node.Nodes) is { } found) return found;
        }
        return null;
    }

    /// <summary>
    /// その行のファイル（フォルダなら中の最初のファイル）のページを右に並べ、行のしおりのページまで動かして枠で囲む。
    /// </summary>
    private void ShowPagesOf(TreeNode node, bool scroll)
    {
        var target = node.Tag is Bookmark { Kind: BookmarkKind.Folder } ? FirstFile(node.Nodes) : node;
        if (target?.Tag is not Bookmark { Part: >= 0 } bookmark)
        {
            // 目次のように、元のファイルのページを指さない行。何も変わらないと理由が分からないので、説明を出す
            _strip.SetCurrent(-1);
            if (target?.Tag is Bookmark { Kind: BookmarkKind.Contents }) ShowContents();
            return;
        }
        ShowPagesNote("", warning: false);
        if (bookmark.Part != _shownPart)
        {
            _shownPart = bookmark.Part;
            _strip.BookmarksEnabled = true;
            _strip.ZoomEnabled = _zoom is not null;
            var file = FileNodeOf(target);
            _pagesTitle.Text = $"{(file?.Tag as Bookmark)?.Title} のページ（番号は木と同じ、結合した PDF でのページ。右クリックでしおりを付ける）";
            if (_pagesOf(bookmark.Part) is { } pages)
            {
                // ページの下の番号を、木の括弧の中と同じ「結合した PDF でのページ」にそろえる。
                // ファイルのしおりのページが、そのファイルで使う最初のページにあたり、そこから使う順に続く
                int start = (file?.Tag as Bookmark)?.Page ?? 0;
                _mergedNumbers = [];
                for (int k = 0; k < pages.Pages.Count; k++) _mergedNumbers.TryAdd(pages.Pages[k], start + k + 1);
                _strip.PageLabel = page => _mergedNumbers.TryGetValue(page, out int number) ? $"{number}" : "-";
                _strip.ShowDocument(pages.Pdf, pages.Pages, pages.Rotations, interactive: false, hint: "");
            }
            else
                _strip.ShowMessage("ページを表示できません");
        }
        _strip.SetCurrent(bookmark.SourcePage);
        if (scroll && bookmark.SourcePage >= 0) _strip.ScrollToPage(bookmark.SourcePage);
    }

    /// <summary>
    /// 結合するときに作る目次ページを、今の名前で作って右に並べる（名前を変えたら作り直す）。
    /// 目次ページにはしおりを付けられないので、右クリックの項目は出さない。
    /// </summary>
    private void ShowContents()
    {
        _shownPart = ContentsPart;
        _strip.BookmarksEnabled = false;
        _strip.ZoomEnabled = false; // 作る前の目次は、拡大の画面では開けない
        _pagesTitle.Text = "目次のページ（結合するときに作る。番号は結合した PDF でのページ）";
        _mergedNumbers = [];
        if (_contents?.Invoke(_choices) is { } pdf)
        {
            for (int i = 0; i < pdf.PageCount; i++) _mergedNumbers[i] = i + 1;
            _strip.PageLabel = page => $"{page + 1}";
            _strip.ShowDocument(pdf, Enumerable.Range(0, pdf.PageCount), FileEntry.NoRotations, interactive: false, hint: "");
            _strip.SetCurrent(0);
            ShowPagesNote("", warning: false);
        }
        else
        {
            _strip.ShowMessage("目次のページを作れませんでした");
        }
    }

    /// <summary>右のページを押したら、そのページを含むしおりを木で選び、押したページに枠を付ける（ページは動かさない）。</summary>
    private void SelectBookmarkAt(int page)
    {
        if (_shownPart == ContentsPart)
        {
            if (Find(_tree.Nodes, b => b.Kind == BookmarkKind.Contents) is { } contents) SelectQuietly(contents);
            _strip.SetCurrent(page);
            return;
        }
        if (_shownPart < 0 || !_mergedNumbers.TryGetValue(page, out int number)) return; // 結合しないページ
        var roots = _tree.Nodes.Cast<TreeNode>().Select(n => n.Tag).OfType<Bookmark>().ToList();
        if (Bookmarks.Covering(roots, _shownPart, number - 1) is not { } bookmark) return;
        if (Find(_tree.Nodes, b => ReferenceEquals(b, bookmark)) is not { } node) return;
        SelectQuietly(node);
        _strip.SetCurrent(page);
    }

    /// <summary>木の行を選ぶ。選び直しで右のページを、しおりの始まりへ動かさない。</summary>
    private void SelectQuietly(TreeNode node)
    {
        _rebuilding = true;
        try
        {
            _tree.SelectedNode = node;
            node.EnsureVisible();
        }
        finally
        {
            _rebuilding = false;
        }
    }

    /// <summary>右のページの下に説明を出す。warning ならできなかったことの説明として赤で出す。空なら消す。</summary>
    private void ShowPagesNote(string text, bool warning)
    {
        _pagesNote.ForeColor = warning ? Theme.Danger : SystemColors.GrayText;
        _pagesNote.Text = text;
    }

    /// <summary>ファイルのしおりの下での段（1 が一番上）。</summary>
    private static int LevelOf(TreeNode node)
    {
        int level = 0;
        for (var n = node; n is not null && n.Tag is not Bookmark { Kind: BookmarkKind.File }; n = n.Parent) level++;
        return level;
    }

    private static bool IsMovable(TreeNode? node) => node?.Tag is Bookmark { Kind: BookmarkKind.Original or BookmarkKind.Added };

    /// <summary>木を作り直しても同じしおりを指す、行の作り。</summary>
    private static string KeyOf(Bookmark b) => $"{b.Kind}|{b.Part}|{b.OutlinePath}|{b.AddedIndex}";

    private bool IsChosen(TreeNode node) =>
        IsMovable(node) && (node == _tree.SelectedNode || _marked.Contains(KeyOf((Bookmark)node.Tag!)));

    /// <summary>
    /// 操作の対象になる行（選んでいる行と、一緒に選んだ行）を上から順に。
    /// 親と子を両方選んでいたら親だけにする（子は親に付いて動くので）。
    /// </summary>
    private List<TreeNode> Chosen()
    {
        var chosen = All(_tree.Nodes).Where(IsChosen).ToList();
        var set = chosen.ToHashSet();
        return chosen.Where(n => !AncestorsOf(n).Any(set.Contains)).ToList();
    }

    private static IEnumerable<TreeNode> All(TreeNodeCollection nodes)
    {
        foreach (TreeNode node in nodes)
            foreach (var n in WithDescendants(node)) yield return n;
    }

    private static IEnumerable<TreeNode> AncestorsOf(TreeNode node)
    {
        for (var n = node.Parent; n is not null; n = n.Parent) yield return n;
    }

    /// <summary>
    /// Ctrl を押しながらのクリックは、その行を一緒に選ぶ（もう選んでいれば外す）。
    /// Shift を押しながらのクリックは、選んでいる行からその行までを一緒に選ぶ。何も押さなければ 1 行だけにする。
    /// </summary>
    private void OnLeftDown(TreeNode? hit)
    {
        if (hit is null) return;
        var keys = ModifierKeys;
        var current = _tree.SelectedNode;
        if (keys.HasFlag(Keys.Control) && IsMovable(hit))
        {
            if (IsMovable(current)) _marked.Add(KeyOf((Bookmark)current!.Tag!));
            string key = KeyOf((Bookmark)hit.Tag!);
            if (hit != current && _marked.Remove(key)) _cancelSelect = true; // 印を外すだけで、選び直さない
            else _marked.Add(key);
            _keepMarks = true;
        }
        else if (keys.HasFlag(Keys.Shift) && IsMovable(hit) && IsMovable(current))
        {
            var visible = All(_tree.Nodes).Where(n => n.IsVisible || AncestorsOf(n).All(a => a.IsExpanded)).ToList();
            int from = visible.IndexOf(current!), to = visible.IndexOf(hit);
            if (from >= 0 && to >= 0)
            {
                foreach (var n in visible.Skip(Math.Min(from, to)).Take(Math.Abs(to - from) + 1).Where(IsMovable))
                    _marked.Add(KeyOf((Bookmark)n.Tag!));
            }
            _keepMarks = true;
        }
        else
        {
            _marked.Clear();
        }
        ApplyMarks();
        UpdateActions();
    }

    /// <summary>一緒に選んだ行に、選んでいる行と同じ色を付ける（TreeView は 1 行しか選べないので、色で見せる）。</summary>
    private void ApplyMarks()
    {
        foreach (var node in All(_tree.Nodes))
        {
            if (node.Tag is not Bookmark bookmark) continue;
            bool marked = node != _tree.SelectedNode && _marked.Contains(KeyOf(bookmark));
            node.BackColor = marked ? SystemColors.Highlight : Color.Empty;
            node.ForeColor = marked ? SystemColors.HighlightText : Color.Empty;
        }
    }

    /// <summary>
    /// 段を変えられるか。上げるのは、どれも一番上の段でないとき。
    /// 下げるのは、どれもすぐ上に同じ段のしおりがあるとき（その下に入る。並んだ兄弟を選んでいれば、そろって同じしおりの下に入る）。
    /// </summary>
    private static bool CanMove(IReadOnlyList<TreeNode> nodes, int delta) =>
        nodes.Count > 0 && nodes.All(n => IsMovable(n) && (delta < 0 ? LevelOf(n) > 1 : n.PrevNode is not null));

    /// <summary>
    /// しおりの段を 1 つ上げる（delta = -1）か下げる（+1）。下のしおりも一緒に動かす。
    /// 段を書き込むのは、動かしたしおりと、下にある「段が書き込み済みのしおり・足したしおり」だけ
    /// （何もしていない元のしおりは、親に付いて動くので書き込まない。書き込むと手直しの件数が増えて見える）。
    /// </summary>
    private void MoveLevel(IReadOnlyList<TreeNode> nodes, int delta)
    {
        if (!CanMove(nodes, delta)) return;
        EndRename(commit: true);
        var primary = _tree.SelectedNode?.Tag as Bookmark;
        foreach (var node in nodes)
            MoveOne(node, delta);
        // 一緒に選んだ行は、行の作りで覚えているので、作り直したあとも選んだまま
        Rebuild(primary is null ? null : b => KeyOf(b) == KeyOf(primary));
    }

    private void MoveOne(TreeNode node, int delta)
    {
        foreach (var n in WithDescendants(node))
        {
            int level = LevelOf(n) + delta;
            switch (n.Tag)
            {
                case Bookmark { Kind: BookmarkKind.Original, OutlinePath: { } path } original:
                    var levels = _choices.LevelsOf(original.Part);
                    if (ReferenceEquals(n, node) || levels.ContainsKey(path)) levels[path] = level;
                    break;
                case Bookmark { Kind: BookmarkKind.Added } added:
                    var list = _choices.Added[added.Part];
                    list[added.AddedIndex] = list[added.AddedIndex] with { Level = level };
                    break;
            }
        }
    }

    private static IEnumerable<TreeNode> WithDescendants(TreeNode node)
    {
        yield return node;
        foreach (TreeNode child in node.Nodes)
            foreach (var n in WithDescendants(child)) yield return n;
    }

    /// <summary>選んだ行に合わせて、操作のボタンを押せるかどうかを決める。</summary>
    private void UpdateActions()
    {
        var chosen = Chosen();
        _promote.Enabled = CanMove(chosen, -1);
        _demote.Enabled = CanMove(chosen, +1);
        _rename.Enabled = IsEditable(_tree.SelectedNode) && chosen.Count <= 1; // 名前は 1 件ずつ
        _remove.Enabled = chosen.Count > 0;
        _remove.Text = chosen.Count > 1 ? $"外す（{chosen.Count} 件）" : "外す";
    }

    private static TreeNode? FileNodeOf(TreeNode? node)
    {
        while (node is not null && node.Tag is not Bookmark { Kind: BookmarkKind.File }) node = node.Parent;
        return node;
    }

    /// <summary>右のページにしおりを足す。すぐ木に出し、その行で名前を書けるようにする。</summary>
    private void AddAt(int page)
    {
        if (_shownPart < 0) return;
        EndRename(commit: true);
        if (_pagesOf(_shownPart) is not { } pages || !pages.Pages.Contains(page))
        {
            System.Media.SystemSounds.Beep.Play();
            ShowPagesNote($"元のファイルの {page + 1} ページは結合しないので、しおりを付けられません（メインの画面で結合に戻せます）。", warning: true);
            return;
        }
        ShowPagesNote("", warning: false);
        if (!_choices.Added.TryGetValue(_shownPart, out var list)) _choices.Added[_shownPart] = list = [];
        // ページ順に並べる。同じページなら、足した順に後ろへ
        int index = list.FindLastIndex(b => b.Page <= page) + 1;
        // 仮の名前は、ページの下に出ている番号（結合した PDF でのページ）にそろえる
        _marked.Clear(); // 足したしおりの番号がずれるので、一緒に選んだ印は消す
        list.Insert(index, new AddedBookmark($"{_mergedNumbers.GetValueOrDefault(page, page + 1)} ページ", page));
        int part = _shownPart;
        Rebuild(b => b.Kind == BookmarkKind.Added && b.Part == part && b.AddedIndex == index);
        // 右クリックのメニューから呼ばれると、この後でメニューが閉じて入力先が元の部品へ戻り、入力欄がすぐ閉じてしまう。
        // メニューが閉じきってから入力欄を開く
        if (_tree.SelectedNode is { Tag: Bookmark { Kind: BookmarkKind.Added } } added) BeginInvoke(() => BeginRename(added));
    }

    /// <summary>設定ファイルの段数を、選べる値にそろえる（負や、選べない段数は「すべて」）。</summary>
    public static int? DepthFrom(int saved) => Array.Exists(Depths, d => d.Depth == saved) ? saved : null;

    /// <summary>既定（すべて・フォルダごとにまとめる・ファイルごとの指定なし）と違う選び方の説明。既定のままなら空。</summary>
    public static string Describe(int? depth, bool folders, bool perFile)
    {
        var parts = new List<string>();
        if (depth == 0) parts.Add("元のしおりなし");
        else if (depth is not null) parts.Add($"元のしおり {depth} 段まで");
        if (!folders) parts.Add("フォルダでまとめない");
        if (perFile) parts.Add("ファイルごとの指定あり");
        return string.Join("・", parts);
    }

    /// <summary>ダイアログで決めた選び方。</summary>
    public BookmarkChoices Choices => _choices;

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _icons.Dispose();
            _menu.Dispose();
            _toolTip.Dispose();
        }
        base.Dispose(disposing);
    }

    private static bool IsEditable(TreeNode? node) =>
        node?.Tag is Bookmark { Kind: BookmarkKind.File or BookmarkKind.Folder or BookmarkKind.Original or BookmarkKind.Added };

    /// <summary>元のしおり（その下のしおりごと）か、足したしおりを外す。外したあとは、上の行を選んでおく。</summary>
    private void Drop(IReadOnlyList<TreeNode> nodes)
    {
        if (nodes.Count == 0) return;
        EndRename(commit: true);
        var bookmarks = nodes.Select(n => (Bookmark)n.Tag!).ToList();
        foreach (var original in bookmarks.Where(b => b.Kind == BookmarkKind.Original))
            _choices.EditsOf(original.Part)[original.OutlinePath!] = null;
        // 足したしおりは番号で指しているので、後ろから外す（前から外すと、後ろの番号がずれる）
        foreach (var added in bookmarks.Where(b => b.Kind == BookmarkKind.Added).OrderByDescending(b => b.AddedIndex))
            _choices.Added[added.Part].RemoveAt(added.AddedIndex);
        _marked.Clear(); // 足したしおりの番号が変わるので、一緒に選んだ印は消す
        _tree.SelectedNode = nodes[0].Parent;
        Rebuild();
    }

    /// <summary>行の上に入力欄を重ねて、名前を変え始める。入力欄には、ページ番号を外した名前だけを入れる。</summary>
    private void BeginRename(TreeNode node)
    {
        // 後回しにして呼ばれたときは、その間に木が作り直されて、行がもう木にないことがある
        if (!IsEditable(node) || node.TreeView is null || IsDisposed) return;
        EndRename(commit: true);
        node.EnsureVisible();
        var bounds = node.Bounds;
        int width = Math.Max(LogicalToDeviceUnits(80),
            Math.Min(Math.Max(bounds.Width + LogicalToDeviceUnits(16), LogicalToDeviceUnits(240)), _tree.ClientSize.Width - bounds.X));
        // 文字の大きさは指定しない。木から受け継がせる（指定すると、表示倍率の調整がずれて小さく出る。一覧の入力欄と同じ）
        var editor = new TextBox
        {
            Text = ((Bookmark)node.Tag!).Title,
            Tag = node,
            BorderStyle = BorderStyle.FixedSingle,
            AccessibleName = "しおりの名前",
        };
        // ほかの場所を押したときは、書きかけの名前で確定する
        editor.LostFocus += (_, _) =>
        {
            if (ReferenceEquals(_editor, editor)) EndRename(commit: true);
        };
        _editor = editor;
        _tree.Controls.Add(editor);
        // 1 行の入力欄の高さは文字で決まるので、木に入れて文字が決まってから、行の上下の中央に置く
        editor.SetBounds(bounds.X, bounds.Y + (bounds.Height - editor.Height) / 2, width, editor.Height);
        editor.BringToFront();
        editor.Focus();
        editor.SelectAll();
    }

    /// <summary>入力欄を閉じる。commit なら名前を選び方に移し、木を作り直す（ページ番号を付け直すため）。</summary>
    private void EndRename(bool commit)
    {
        if (_editor is not { } editor) return;
        _editor = null;
        var node = (TreeNode)editor.Tag!;
        string name = editor.Text.Trim();
        _tree.Controls.Remove(editor);
        editor.Dispose();
        _tree.Focus();
        if (commit && node.Tag is Bookmark bookmark && Rename(bookmark, name)) Rebuild();
    }

    // 入力中の Enter と Esc は、OK とキャンセルのボタンより先に入力欄で受け取る
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (_editor is not null && keyData is Keys.Enter or Keys.Escape)
        {
            EndRename(commit: keyData == Keys.Enter);
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    /// <returns>選び方が変わったら true。</returns>
    private bool Rename(Bookmark bookmark, string name)
    {
        if (name == bookmark.Title) return false; // 変えていない
        if (bookmark.Kind == BookmarkKind.File)
        {
            _choices.Titles[bookmark.Part] = name; // 空なら、結合するときに元のファイル名に戻る
        }
        else if (bookmark.Kind == BookmarkKind.Added)
        {
            if (name.Length == 0) return false; // 足したしおりには戻す名前がないので、空にはしない（外すときは Delete）
            var list = _choices.Added[bookmark.Part];
            list[bookmark.AddedIndex] = list[bookmark.AddedIndex] with { Title = name };
        }
        else if (bookmark.Kind == BookmarkKind.Original)
        {
            // 空にしたときは、元の名前に戻す
            var edits = _choices.EditsOf(bookmark.Part);
            if (name.Length == 0) edits.Remove(bookmark.OutlinePath!);
            else edits[bookmark.OutlinePath!] = name;
        }
        else if (bookmark.FolderKey is { } key)
        {
            // 空にするか元の名前に戻したときは、付け直しをやめる
            if (name.Length == 0 || name == OriginalFolderName(key)) _choices.FolderNames.Remove(key);
            else _choices.FolderNames[key] = name;
        }
        return true;
    }

    /// <summary>フォルダの Key（完全なパス）から、元の名前を求める。ドライブ直下のように名前がなければパスそのもの（BookmarkFolders と同じ）。</summary>
    private static string OriginalFolderName(string key) => Path.GetFileName(key) is { Length: > 0 } name ? name : key;

    private void OnMenuOpening(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _menu.Items.Clear();
        // マウスで開いたときは押した場所の行、キーボード（Shift+F10）で開いたときは選んでいる行
        var node = _menuByMouse ? _menuTarget : _tree.SelectedNode;
        _menuByMouse = false;
        if (!IsEditable(node))
        {
            e.Cancel = true;
            return;
        }
        var bookmark = (Bookmark)node!.Tag!;
        e.Cancel = false; // 項目が空のメニューは、開く前から取りやめの印が付いている。ここで項目を足すので外す
        _menu.Items.Add(new ToolStripMenuItem("名前を変える", null, (_, _) => BeginInvoke(() => BeginRename(node))) { ShortcutKeyDisplayString = "F2" });
        if (bookmark.Kind is BookmarkKind.Original or BookmarkKind.Added)
        {
            var chosen = Chosen();
            _menu.Items.Add(new ToolStripMenuItem("1 段上げる", null, (_, _) => MoveLevel(chosen, -1))
            {
                ShortcutKeyDisplayString = "Alt+Shift+←",
                Enabled = CanMove(chosen, -1),
            });
            _menu.Items.Add(new ToolStripMenuItem("1 段下げる", null, (_, _) => MoveLevel(chosen, +1))
            {
                ShortcutKeyDisplayString = "Alt+Shift+→",
                Enabled = CanMove(chosen, +1),
            });
        }
        if (bookmark.Kind == BookmarkKind.Added)
        {
            _menu.Items.Add(new ToolStripMenuItem("このしおりを外す", null, (_, _) => Drop(Chosen())) { ShortcutKeyDisplayString = "Delete" });
            return;
        }
        if (bookmark.Kind == BookmarkKind.Original)
        {
            _menu.Items.Add(new ToolStripMenuItem("このしおりを外す（下のしおりも外れます）", null, (_, _) => Drop(Chosen()))
            {
                ShortcutKeyDisplayString = "Delete",
            });
            return;
        }
        if (bookmark.Kind == BookmarkKind.Folder && bookmark.FolderKey is { } key && _choices.FolderNames.ContainsKey(key))
        {
            _menu.Items.Add("元の名前に戻す", null, (_, _) =>
            {
                _choices.FolderNames.Remove(key);
                Rebuild();
            });
        }
        if (bookmark.Kind != BookmarkKind.File) return;

        if (_choices.OutlineEdits.GetValueOrDefault(bookmark.Part) is { Count: > 0 }
            || _choices.OutlineLevels.GetValueOrDefault(bookmark.Part) is { Count: > 0 })
        {
            _menu.Items.Add("元のしおりの手直しを取り消す", null, (_, _) =>
            {
                _choices.OutlineEdits.Remove(bookmark.Part);
                _choices.OutlineLevels.Remove(bookmark.Part);
                Rebuild();
            });
        }

        _menu.Items.Add(new ToolStripSeparator());
        _menu.Items.Add(new ToolStripMenuItem("このファイルの元のしおり") { Enabled = false });
        int? current = _choices.FileDepths.GetValueOrDefault(bookmark.Part);
        foreach (var (text, depth) in FileDepthChoices)
        {
            _menu.Items.Add(new ToolStripMenuItem(text, null, (_, _) =>
            {
                _choices.FileDepths[bookmark.Part] = depth;
                Rebuild();
            })
            { Checked = current == depth });
        }
    }

    /// <summary>ファイルの行に添える、そのファイルだけの指定（段数と、手直しと足したしおりの件数）。</summary>
    private string OwnText(int part)
    {
        var parts = new List<string>();
        if (FileDepthText(_choices.FileDepths.GetValueOrDefault(part)) is { Length: > 0 } depth) parts.Add(depth);
        int edited = (_choices.OutlineEdits.GetValueOrDefault(part)?.Count ?? 0) + (_choices.OutlineLevels.GetValueOrDefault(part)?.Count ?? 0);
        if (edited > 0) parts.Add($"手直し {edited} 件");
        if (_choices.Added.GetValueOrDefault(part) is { Count: > 0 } added) parts.Add($"足したしおり {added.Count} 件");
        return string.Join("・", parts);
    }

    private static string FileDepthText(int? depth) =>
        depth switch
        {
            null => "",
            0 => "元のしおりなし",
            MergePart.AllLevels => "元のしおりはすべて",
            _ => $"元のしおり {depth} 段まで",
        };

    /// <param name="select">作り直したあとに選ぶ行。null なら、前に選んでいた行を選び直す。</param>
    private void Rebuild(Func<Bookmark, bool>? select = null)
    {
        _rebuilding = true;
        // 作り直したあとも、選んでいたファイルやフォルダを選んだままにする
        var selected = _tree.SelectedNode?.Tag as Bookmark;
        _tree.BeginUpdate();
        _tree.Nodes.Clear();
        var bookmarks = _build(_choices);
        // フォルダの段があるときだけアイコンの一覧を付ける。付けると、アイコンのない行にも空のアイコンの幅が取られるので
        var icons = HasFolder(bookmarks) ? _icons : null;
        if (!ReferenceEquals(_tree.ImageList, icons))
        {
            _tree.ImageList = icons;
            FitRows();
        }
        TreeNode? reselect = null;
        Add(_tree.Nodes, bookmarks, selected, ref reselect);
        if (select is not null) reselect = Find(_tree.Nodes, select) ?? reselect;
        if (bookmarks.Count == 0) _tree.Nodes.Add("（結合できるファイルがありません）").ForeColor = SystemColors.GrayText;
        _tree.ExpandAll();
        if (reselect is not null)
        {
            _tree.SelectedNode = reselect;
            reselect.EnsureVisible();
        }
        else if (_tree.Nodes.Count > 0)
        {
            _tree.Nodes[0].EnsureVisible();
        }
        _tree.EndUpdate();
        _rebuilding = false;
        ApplyMarks();
        UpdateActions();
        _strip.Invalidate(); // ページの下の、付けたしおりの名前を描き直す
        // 目次ページを並べているときは、名前の変更を目次にも映す
        if (_shownPart == ContentsPart && select is null) ShowContents();
        // 選んだ行のファイルが変わったとき（足したばかりの行など）は、右のページも合わせる
        if (select is not null && reselect is not null) ShowPagesOf(reselect, scroll: true);
    }

    private static bool HasFolder(IEnumerable<Bookmark> nodes) =>
        nodes.Any(n => n.Kind == BookmarkKind.Folder || HasFolder(n.Children));

    private static TreeNode? Find(TreeNodeCollection nodes, Func<Bookmark, bool> match)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Tag is Bookmark bookmark && match(bookmark)) return node;
            if (Find(node.Nodes, match) is { } found) return found;
        }
        return null;
    }

    private void Add(TreeNodeCollection target, IReadOnlyList<Bookmark> nodes, Bookmark? selected, ref TreeNode? reselect)
    {
        foreach (var bookmark in nodes)
        {
            string own = bookmark.Kind == BookmarkKind.File ? OwnText(bookmark.Part) : "";
            var node = target.Add(own.Length > 0 ? $"{bookmark.Title}（{bookmark.Page + 1}）　［{own}］" : $"{bookmark.Title}（{bookmark.Page + 1}）");
            node.Tag = bookmark;
            node.ToolTipText = bookmark.Kind switch
            {
                BookmarkKind.Contents => "目次のページ",
                BookmarkKind.Folder => "フォルダ。名前は F2 で変えられます",
                BookmarkKind.File => "ファイル。名前は F2 で変えられます（一覧の「しおり名」も変わります）",
                BookmarkKind.Added => "ページを見て足したしおり。名前は F2 で変えられ、Delete で外せます",
                _ => "元のファイルのしおり。名前は F2 で変えられ、Delete で外せます（下のしおりも外れます）",
            };
            node.ImageIndex = node.SelectedImageIndex = bookmark.Kind == BookmarkKind.Folder ? FolderIcon : NoIcon;
            if (selected is not null && bookmark.Kind == selected.Kind && bookmark.Kind is not BookmarkKind.Contents
                && bookmark.Part == selected.Part && bookmark.FolderKey == selected.FolderKey && bookmark.OutlinePath == selected.OutlinePath
                && bookmark.AddedIndex == selected.AddedIndex)
            {
                reselect = node;
            }
            Add(node.Nodes, bookmark.Children, selected, ref reselect);
        }
    }
}
