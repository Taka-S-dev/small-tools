using DocsToPdf.Core;

namespace DocsToPdf.Ui;

/// <summary>Excel のどのシートを出すかと、幅を 1 ページに収めるかを決める。</summary>
internal sealed class ExcelSheetsDialog : Form
{
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    private readonly CheckedListBox _sheets = new() { CheckOnClick = true, IntegralHeight = false, Dock = DockStyle.Fill };
    private readonly CheckBox _fitWidth = new() { AutoSize = true, Text = "すべての列を 1 ページの幅に収める（横に分かれて出るのを防ぐ）" };
    private readonly Button _ok = DialogButton("OK", DialogResult.OK);
    private readonly IReadOnlyList<SheetInfo> _all;

    public ExcelSheetsDialog(string fileName, IReadOnlyList<SheetInfo> sheets, bool fitWidth)
    {
        _all = sheets;
        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = fileName;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(440, 380);
        MinimumSize = new Size(360, 300);

        foreach (var sheet in sheets)
        {
            string pages = sheet.Included && sheet.Pages > 0 ? $"（{sheet.Pages} ページ）"
                : sheet.Included ? "（出力するものがありません）" : "";
            _sheets.Items.Add(sheet.Name + pages, sheet.Included);
        }
        _fitWidth.Checked = fitWidth;

        var selectAll = DialogButton("すべて選ぶ");
        var selectNone = DialogButton("すべて外す");
        selectAll.Click += (_, _) => SetAll(true);
        selectNone.Click += (_, _) => SetAll(false);
        var cancel = DialogButton("キャンセル", DialogResult.Cancel);

        var pick = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 0, 6) };
        pick.Controls.AddRange([selectAll, selectNone]);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 10, 0, 0) };
        buttons.Controls.AddRange([cancel, _ok]);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
        // 列はウィンドウの幅に合わせる（中身に合わせると、右寄せの OK・キャンセルが外へ押し出される）
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(new Label { Text = "PDF にするシート（非表示のシートは出ません）", AutoSize = true, Margin = new Padding(0, 0, 0, 6) });
        root.Controls.Add(_sheets);
        root.Controls.Add(pick);
        root.Controls.Add(_fitWidth);
        root.Controls.Add(buttons);
        Controls.Add(root);

        AcceptButton = _ok;
        CancelButton = cancel;
        // ItemCheck はチェックが変わる前に来るので、変わった後に数え直す
        _sheets.ItemCheck += (_, _) => BeginInvoke(UpdateOk);
        UpdateOk();
        ResumeLayout(false);
        PerformLayout();
    }

    /// <summary>選んだシート名。すべて選んだときは null（シートが増えても全部出す、の意味）。</summary>
    public IReadOnlyList<string>? SelectedSheets
    {
        get
        {
            if (_sheets.CheckedIndices.Count == _all.Count) return null;
            return _sheets.CheckedIndices.Cast<int>().Select(i => _all[i].Name).ToList();
        }
    }

    public bool FitWidth => _fitWidth.Checked;

    private void SetAll(bool value)
    {
        for (int i = 0; i < _sheets.Items.Count; i++) _sheets.SetItemChecked(i, value);
    }

    private void UpdateOk() => _ok.Enabled = _sheets.CheckedIndices.Count > 0;

    /// <summary>
    /// 文字に合わせた大きさのボタン。縮むようにしておかないと、表示倍率に合わせて拡大された既定の大きさが残る。
    /// </summary>
    internal static Button DialogButton(string text, DialogResult result = DialogResult.None) => new()
    {
        Text = text,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        MinimumSize = new Size(88, 0),
        Padding = new Padding(8, 3, 8, 3),
        DialogResult = result,
    };
}

/// <summary>パスワード付きファイルのパスワードを尋ねる。</summary>
internal sealed class PasswordDialog : Form
{
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Theme.ApplyTitleBar(this);
    }

    private readonly TextBox _password = new() { UseSystemPasswordChar = true, Dock = DockStyle.Fill };

    public PasswordDialog(string fileName, string? message)
    {
        SuspendLayout(); // 表示倍率の調整を ResumeLayout まで待たせる（MainForm と同じ）
        Text = "パスワード";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;

        var show = new CheckBox { Text = "パスワードを表示する", AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
        show.CheckedChanged += (_, _) => _password.UseSystemPasswordChar = !show.Checked;
        var ok = ExcelSheetsDialog.DialogButton("OK", DialogResult.OK);
        var cancel = ExcelSheetsDialog.DialogButton("キャンセル", DialogResult.Cancel);
        var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 12, 0, 0) };
        buttons.Controls.AddRange([cancel, ok]);

        var root = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Padding = new Padding(12), MinimumSize = new Size(380, 0) };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.Controls.Add(new Label
        {
            Text = $"{fileName}\r\n{message ?? "このファイルを開くためのパスワードを入力してください。"}",
            AutoSize = true,
            MaximumSize = new Size(420, 0),
            Margin = new Padding(0, 0, 0, 8),
        });
        root.Controls.Add(_password);
        root.Controls.Add(show);
        root.Controls.Add(buttons);
        Controls.Add(root);

        AcceptButton = ok;
        CancelButton = cancel;
        ResumeLayout(false);
        PerformLayout();
    }

    public string Password => _password.Text;
}
