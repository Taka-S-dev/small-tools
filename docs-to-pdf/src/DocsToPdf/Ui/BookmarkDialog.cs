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

    /// <summary>そのファイルの手直しを書き込む先（なければ作る）。</summary>
    public Dictionary<string, string?> EditsOf(int part) =>
        OutlineEdits.TryGetValue(part, out var edits) ? edits : OutlineEdits[part] = [];
}

/// <summary>
/// 結合した PDF に付くしおりを、作る前に木の形で見せる。ファイル・フォルダ・元のしおりの名前はその場で変えられ、
/// 元のしおりは 1 件ずつ外せる。元のしおりを何段まで付けるかは、全体とファイルごとに選べる。
/// 変えるたびに木を作り直す（結合と同じ決まりで求める）。
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
    private readonly ContextMenuStrip _menu = new();
    private TreeNode? _menuTarget;
    private bool _menuByMouse;
    // フォルダの段にだけフォルダのアイコンを付ける（ほかの行は、幅をそろえるための空の絵）
    private readonly ImageList _icons = new() { ColorDepth = ColorDepth.Depth32Bit };
    private const int NoIcon = 0, FolderIcon = 1;

    /// <param name="build">選び方 → できあがるしおり。</param>
    /// <param name="choices">今の選び方。ダイアログの中で書き換えるので、呼ぶ側は写しを渡す。</param>
    /// <param name="note">一覧のうち、しおりに入らないファイルがあるときの説明（変換中や、結合できないもの）。なければ空。</param>
    public BookmarkDialog(Func<BookmarkChoices, IReadOnlyList<Bookmark>> build, BookmarkChoices choices, string note)
    {
        _build = build;
        _choices = choices;
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
        ClientSize = new Size(480, 560);
        MinimumSize = new Size(360, 340);

        _icons.ImageSize = SystemInformation.SmallIconSize;
        _icons.Images.Add(new Bitmap(_icons.ImageSize.Width, _icons.ImageSize.Height));
        using (var folder = ShellIcons.ForFolder()) _icons.Images.Add(folder?.ToBitmap() ?? new Bitmap(_icons.ImageSize.Width, _icons.ImageSize.Height));
        _tree.ImageList = _icons;
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
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(depthRow);
        root.Controls.Add(_folders);
        root.Controls.Add(new Label { Text = "できあがるしおり（括弧の中は、結合した PDF でのページ）", AutoSize = true, Margin = new Padding(0, 6, 0, 4) });
        root.Controls.Add(_tree);
        root.Controls.Add(new Label
        {
            Text = "名前は F2 で変えられ、元のしおりは Delete で外せます。右クリックで、ファイルごとに元のしおりを何段まで付けるかも選べます。",
            AutoSize = true,
            MaximumSize = new Size(456, 0),
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
        _tree.KeyDown += (_, e) =>
        {
            _menuByMouse = false; // この後の Shift+F10 は、選んでいる行のメニューにする
            if (e.KeyCode == Keys.F2 && _tree.SelectedNode is { } node) BeginRename(node);
            else if (e.KeyCode == Keys.Delete && _tree.SelectedNode is { Tag: Bookmark { Kind: BookmarkKind.Original } } original) Drop(original);
        };
        // 右クリックした行を選んでから、メニューを開く（TreeView は右クリックでは選ばない）。
        // 行のないところで右クリックしたときは、前に選んでいた行のメニューを出さない
        _tree.MouseDown += (_, e) =>
        {
            if (e.Button != MouseButtons.Right) return;
            _menuTarget = _tree.GetNodeAt(e.Location);
            _menuByMouse = true;
            if (_menuTarget is not null) _tree.SelectedNode = _menuTarget;
        };
        _menu.Opening += OnMenuOpening;
        ResumeLayout(false);
        PerformLayout();
        Rebuild();
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
        }
        base.Dispose(disposing);
    }

    private static bool IsEditable(TreeNode? node) =>
        node?.Tag is Bookmark { Kind: BookmarkKind.File or BookmarkKind.Folder or BookmarkKind.Original };

    /// <summary>元のしおりを、その下のしおりごと外す。外したあとは、上の行を選んでおく。</summary>
    private void Drop(TreeNode node)
    {
        var bookmark = (Bookmark)node.Tag!;
        _choices.EditsOf(bookmark.Part)[bookmark.OutlinePath!] = null;
        _tree.SelectedNode = node.Parent;
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
        if (bookmark.Kind == BookmarkKind.Original)
        {
            _menu.Items.Add(new ToolStripMenuItem("このしおりを外す（下のしおりも外れます）", null, (_, _) => Drop(node))
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

        if (_choices.OutlineEdits.GetValueOrDefault(bookmark.Part) is { Count: > 0 })
        {
            _menu.Items.Add("元のしおりの手直しを取り消す", null, (_, _) =>
            {
                _choices.OutlineEdits.Remove(bookmark.Part);
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

    /// <summary>ファイルの行に添える、そのファイルだけの指定（段数と、元のしおりの手直しの件数）。</summary>
    private string OwnText(int part)
    {
        var parts = new List<string>();
        if (FileDepthText(_choices.FileDepths.GetValueOrDefault(part)) is { Length: > 0 } depth) parts.Add(depth);
        if (_choices.OutlineEdits.GetValueOrDefault(part) is { Count: > 0 } edits) parts.Add($"手直し {edits.Count} 件");
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

    private void Rebuild()
    {
        // 作り直したあとも、選んでいたファイルやフォルダを選んだままにする
        var selected = _tree.SelectedNode?.Tag as Bookmark;
        _tree.BeginUpdate();
        _tree.Nodes.Clear();
        var bookmarks = _build(_choices);
        TreeNode? reselect = null;
        Add(_tree.Nodes, bookmarks, selected, ref reselect);
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
                _ => "元のファイルのしおり。名前は F2 で変えられ、Delete で外せます（下のしおりも外れます）",
            };
            node.ImageIndex = node.SelectedImageIndex = bookmark.Kind == BookmarkKind.Folder ? FolderIcon : NoIcon;
            if (selected is not null && bookmark.Kind == selected.Kind && bookmark.Kind is not BookmarkKind.Contents
                && bookmark.Part == selected.Part && bookmark.FolderKey == selected.FolderKey && bookmark.OutlinePath == selected.OutlinePath)
            {
                reselect = node;
            }
            Add(node.Nodes, bookmark.Children, selected, ref reselect);
        }
    }
}
