using System.Media;
using DocsToPdf.Core;

namespace DocsToPdf.Ui;

// 一覧の操作: 追加・削除・並べ替え・元に戻す・その場での編集・ドラッグ＆ドロップ・裏での変換
internal sealed partial class MainForm
{
    private const string RowDragFormat = "DocsToPdf.Rows";
    private const int UndoLimit = 100;

    /// <summary>元に戻すための、ある時点の一覧の写し。並び順と、利用者が決めた値を持つ。</summary>
    private sealed record Snapshot(FileEntry Entry, string Title, string PageRange, IReadOnlyList<string>? Sheets, bool FitWidth,
        IReadOnlyDictionary<int, int> Rotations, int? OutlineDepth, OutlineChanges? OutlineChanges, PageBookmarks? PageBookmarks);

    /// <summary>元に戻す 1 回分。行の写しと、しおりの確認で決めたもの（フォルダの段の名前、全体の段数、フォルダごとにまとめるか）。</summary>
    private sealed record ListState(Snapshot[] Rows, IReadOnlyDictionary<string, string> FolderNames, int? OutlineDepth, bool FolderBookmarks);

    private readonly List<ListState> _undoStack = [];
    private readonly List<ListState> _redoStack = [];
    private TextBox? _editor;

    private IEnumerable<ListViewItem> Items => _list.Items.Cast<ListViewItem>();
    private IEnumerable<FileEntry> Entries() => Items.Select(EntryOf);
    private IEnumerable<FileEntry> SelectedEntries() => _list.SelectedItems.Cast<ListViewItem>().Select(EntryOf);
    private static FileEntry EntryOf(ListViewItem item) => (FileEntry)item.Tag!;
    private ListViewItem? ItemOf(FileEntry entry) => Items.FirstOrDefault(i => ReferenceEquals(i.Tag, entry));

    private static string FileDialogFilter
    {
        get
        {
            string all = string.Join(";", FileEntry.SupportedExtensions.Select(e => "*" + e));
            return $"結合できるファイル|{all}|すべてのファイル|*.*";
        }
    }

    private void ChooseFiles()
    {
        if (_busy) return;
        using var dialog = new OpenFileDialog { Multiselect = true, Filter = FileDialogFilter, Title = "結合するファイルを選択" };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            _ = AddPathsAsync(dialog.FileNames, null);
    }

    private void ChooseFolder()
    {
        if (_busy) return;
        using var dialog = new FolderBrowserDialog
        {
            Description = "フォルダの中の対象ファイルをすべて追加します（サブフォルダも含む）",
            UseDescriptionForTitle = true,
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            _ = AddPathsAsync([dialog.SelectedPath], null);
    }

    private void PasteFiles()
    {
        if (_busy || !Clipboard.ContainsFileDropList()) return;
        var paths = Clipboard.GetFileDropList().Cast<string>().ToArray();
        int? after = _list.SelectedIndices.Count > 0 ? _list.SelectedIndices.Cast<int>().Max() + 1 : null;
        _ = AddPathsAsync(paths, after);
    }

    /// <summary>これより多いファイルが一度に見つかったら、本当に全部加えるかを確かめる（Office での変換に時間がかかるため）。</summary>
    private const int ConfirmAddAbove = 200;

    /// <summary>前の追加（フォルダの中を探している途中かもしれない）。続けて落とされたものは、これが終わってから順に加える。</summary>
    private Task _adding = Task.CompletedTask;
    /// <summary>フォルダの中を探している間だけ入る。「中止」と Esc で止める。</summary>
    private CancellationTokenSource? _scan;

    private bool IsScanning => _scan is not null;

    /// <summary>
    /// ファイルやフォルダを加え、すぐに裏で変換を始める。
    /// まとめて渡されたものはパスの順に並べてから insertAt の位置に入れる。
    /// フォルダは画面を止めないよう裏で探し、その間は進み具合を出して中止できるようにする。
    /// </summary>
    private Task AddPathsAsync(IEnumerable<string> paths, int? insertAt)
    {
        string[] list = [.. paths];
        _adding = AddAfterAsync(_adding, list, insertAt);
        return _adding;
    }

    private async Task AddAfterAsync(Task previous, string[] paths, int? insertAt)
    {
        try { await previous; } catch { /* 前の追加の失敗は、そこで知らせ済み */ }
        if (IsDisposed) return;
        try
        {
            InputFiles.Result result;
            if (!paths.Any(Directory.Exists))
            {
                result = InputFiles.Expand(paths); // ファイルだけなら一瞬で済む
            }
            else
            {
                using var scan = new CancellationTokenSource();
                _scan = scan;
                UpdateButtons();
                Cursor = Cursors.AppStarting;
                SetStatus("フォルダの中を探しています...（「中止」か Esc で止められます）");
                var progress = new Progress<(int Scanned, int Found)>(p =>
                {
                    if (ReferenceEquals(_scan, scan))
                        SetStatus($"フォルダの中を探しています... {p.Scanned:N0} 件を調べました（対象 {p.Found:N0} 件）。「中止」か Esc で止められます");
                });
                try
                {
                    result = await Task.Run(() => InputFiles.Expand(paths, progress: progress, ct: scan.Token), scan.Token);
                }
                catch (OperationCanceledException)
                {
                    SetStatus("フォルダの追加を中止しました。");
                    return;
                }
                finally
                {
                    _scan = null;
                    Cursor = Cursors.Default;
                    UpdateButtons();
                }
            }
            if (!IsDisposed) AddFiles(result, insertAt);
        }
        catch (Exception ex)
        {
            ShowError("ファイルを追加できませんでした", ex.Message);
        }
    }

    private void CancelScan()
    {
        if (_scan is not { IsCancellationRequested: false } scan) return;
        scan.Cancel();
        SetStatus("中止しています...");
    }

    private void AddFiles(InputFiles.Result result, int? insertAt)
    {
        if (_busy)
        {
            SetStatus("作成中のため、ファイルを追加できませんでした。終わってからもう一度追加してください。");
            return;
        }
        var known = new HashSet<string>(Entries().Select(e => e.Path), StringComparer.OrdinalIgnoreCase);
        var (files, rejected) = result;
        int unsupported = rejected.Count;

        int fresh = files.Count(f => !known.Contains(f));
        if (fresh > ConfirmAddAbove && !_startup.IsBatch && !ConfirmManyFiles(fresh))
        {
            SetStatus("追加をやめました。");
            return;
        }

        int duplicates = 0;
        var added = new List<FileEntry>();
        foreach (var file in files)
        {
            if (!known.Add(file)) { duplicates++; continue; }
            FileEntry.TryGetKind(file, out var kind);
            added.Add(new FileEntry(file, kind) { Root = result.RootOf(file) });
        }

        if (added.Count > 0)
        {
            PushUndo();
            int index = Math.Clamp(insertAt ?? _list.Items.Count, 0, _list.Items.Count);
            _list.BeginUpdate();
            _list.SelectedItems.Clear();
            var items = added.Select(CreateItem).ToList();
            for (int i = 0; i < items.Count; i++)
                _list.Items.Insert(index + i, items[i]);
            // 1 件だけ追加したときは選んでおき、すぐプレビューが出るようにする
            if (items.Count == 1) items[0].Selected = true;
            _list.EndUpdate();
            items[0].EnsureVisible();
            items[0].Focused = true;
            foreach (var entry in added) Prepare(entry);
        }

        RefreshAll();

        var notes = new List<string>();
        if (duplicates > 0) notes.Add($"追加済み {duplicates} 件");
        if (unsupported > 0) notes.Add($"対象外 {unsupported} 件");
        string suffix = notes.Count > 0 ? $"（{string.Join("、", notes)}は除外）" : "";
        SetStatus(added.Count > 0 ? $"{added.Count} 件追加しました{suffix}" : $"追加できるファイルがありませんでした{suffix}");
    }

    private bool ConfirmManyFiles(int count)
    {
        var add = new TaskDialogButton("すべて追加する");
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = $"{count:N0} 件のファイルが見つかりました",
            Text = "すべて一覧に加えて、PDF への変換を始めます。Word や Excel のファイルが多いと、変換に時間がかかります。",
            Icon = TaskDialogIcon.Information,
            Buttons = { add, TaskDialogButton.Cancel },
            DefaultButton = TaskDialogButton.Cancel,
        };
        return TaskDialog.ShowDialog(this, page) == add;
    }

    private ListViewItem CreateItem(FileEntry entry)
    {
        string extension = Path.GetExtension(entry.Path).ToLowerInvariant();
        if (!_icons.Images.ContainsKey(extension))
        {
            using var icon = ShellIcons.ForExtension(extension);
            if (icon is not null) _icons.Images.Add(extension, icon);
        }
        return new ListViewItem(["", "", "", "", "", ""]) { Tag = entry, ImageKey = extension };
    }

    /// <summary>その行の変換を頼む（頼み直す）。前の依頼の結果は捨てる。</summary>
    private void Prepare(FileEntry entry)
    {
        entry.Pending?.Cancel();
        var pending = new CancellationTokenSource();
        entry.Pending = pending;
        int version = ++entry.Version;
        entry.State = EntryState.Queued;
        entry.Error = null;
        entry.Converted = null;

        var request = new ConvertRequest(entry.Path, entry.Kind, entry.Sheets, entry.FitWidth, entry.Password);
        _ = ConvertAsync(entry, version, request, pending.Token);
        EntryChanged(entry);
    }

    private async Task ConvertAsync(FileEntry entry, int version, ConvertRequest request, CancellationToken token)
    {
        ConvertOutcome outcome;
        try
        {
            outcome = await _service.ConvertAsync(request, onStarted: () => Post(() =>
            {
                if (entry.Version != version || entry.State != EntryState.Queued) return;
                entry.State = EntryState.Converting;
                EntryChanged(entry);
            }), token);
        }
        catch (OperationCanceledException)
        {
            return;
        }
        catch (Exception ex)
        {
            outcome = ConvertOutcome.Failure(ex.Message, null);
        }

        if (IsDisposed || entry.Version != version) return;
        entry.Stamp = outcome.Stamp;
        if (outcome.Pdf is { } pdf)
        {
            entry.Converted = pdf;
            entry.State = EntryState.Ready;
            entry.Error = null;
            // 元のファイルが変わってしおりやページの並びが変わったら、位置やページで当てている手直しは
            // 別の見出しやページに当たるので捨てる
            bool dropped = false;
            if (entry.OutlineChanges is { } changes && !changes.AppliesTo(pdf.Outline))
            {
                entry.OutlineChanges = null;
                dropped = true;
            }
            if (entry.PageBookmarks is { } added && !added.AppliesTo(pdf.PageSizes))
            {
                entry.PageBookmarks = null;
                dropped = true;
            }
            if (dropped)
            {
                _droppedEdits.Add(entry.FileName);
                SetStatus(DroppedEditsNotice());
            }
        }
        else
        {
            entry.State = outcome.NeedsPassword ? EntryState.NeedsPassword : EntryState.Failed;
            entry.Error = outcome.Error;
        }
        EntryChanged(entry);
    }

    /// <summary>別のスレッドから、UI スレッドで処理を行わせる。</summary>
    private void Post(Action action)
    {
        try
        {
            if (IsHandleCreated && !IsDisposed) BeginInvoke(action);
        }
        catch (InvalidOperationException) { /* 閉じている途中 */ }
    }

    /// <summary>一覧に戻ってきた行が、今のファイルの内容で変換済みかを確かめ、必要なら変換し直す。</summary>
    private void EnsureFresh(FileEntry entry)
    {
        if (entry.IsBusy && entry.Pending is { IsCancellationRequested: false }) return;
        var stamp = SourceStamp.Of(entry.Path);
        bool current = entry.State == EntryState.Ready && entry.Converted is { } pdf
            && stamp == entry.Stamp && File.Exists(pdf.PdfPath);
        if (current || (entry.State == EntryState.NeedsPassword && stamp == entry.Stamp)) return;
        Prepare(entry);
    }

    private Task _sourceCheck = Task.CompletedTask;

    /// <summary>
    /// 元のファイルが更新・削除されていないかを見て、変わった行だけ変換し直す。
    /// ファイルを見に行く部分は別のスレッドで行う（応答しないネットワーク上のファイルで画面を固めないため）。
    /// すでに確認中なら、その確認が終わるのを待つだけにする。
    /// </summary>
    private Task CheckSourcesAsync()
    {
        if (!_sourceCheck.IsCompleted) return _sourceCheck;
        return _sourceCheck = CheckSourcesCoreAsync();
    }

    private async Task CheckSourcesCoreAsync()
    {
        var targets = Entries().Where(e => !e.IsBusy).Select(e => (Entry: e, e.Version)).ToList();
        if (targets.Count == 0) return;
        var stamps = await Task.Run(() => targets.Select(t => SourceStamp.Of(t.Entry.Path)).ToList());
        if (IsDisposed) return;

        for (int i = 0; i < targets.Count; i++)
        {
            var (entry, version) = targets[i];
            // 確認している間に変換し直された行や、一覧から外された行は対象外
            if (entry.Version != version || entry.IsBusy || ItemOf(entry) is null) continue;

            var stamp = stamps[i];
            if (stamp is null)
            {
                if (entry.State == EntryState.Failed && entry.Stamp is null) continue;
                entry.Version++;
                entry.State = EntryState.Failed;
                entry.Error = "ファイルが見つかりません。";
                entry.Converted = null;
                entry.Stamp = null;
                EntryChanged(entry);
            }
            else if (stamp != entry.Stamp)
            {
                Prepare(entry);
            }
        }
    }

    private void Retry(IEnumerable<FileEntry> entries)
    {
        if (_busy) return;
        foreach (var entry in entries.Where(e => !e.IsBusy).ToList())
            Prepare(entry);
    }

    private void RetrySelected() => Retry(SelectedEntries());

    private void RemoveSelected()
    {
        if (_busy || _list.SelectedItems.Count == 0) return;
        PushUndo();
        int first = _list.SelectedIndices.Cast<int>().Min();
        _list.BeginUpdate();
        foreach (var item in _list.SelectedItems.Cast<ListViewItem>().ToList())
        {
            Detach(EntryOf(item));
            _list.Items.Remove(item);
        }
        _list.EndUpdate();
        if (_list.Items.Count > 0)
        {
            var next = _list.Items[Math.Min(first, _list.Items.Count - 1)];
            next.Selected = true;
            next.Focused = true;
        }
        RefreshAll();
        SetStatus("リストから外しました（Ctrl+Z で戻せます）。");
    }

    private void ClearAll()
    {
        if (_busy || _list.Items.Count == 0) return;
        PushUndo();
        foreach (var entry in Entries()) Detach(entry);
        _list.Items.Clear();
        RefreshAll();
        SetStatus("リストを空にしました（Ctrl+Z で戻せます）。");
    }

    /// <summary>一覧から外れる行の、まだ始まっていない変換をやめる。</summary>
    private static void Detach(FileEntry entry)
    {
        if (!entry.IsBusy) return;
        entry.Pending?.Cancel();
        entry.Version++;
    }

    private void MoveSelectedBy(int delta)
    {
        if (_busy || _list.SelectedItems.Count == 0) return;
        var indices = _list.SelectedIndices.Cast<int>().Order().ToList();
        if (delta < 0 && indices[0] == 0) return;
        if (delta > 0 && indices[^1] == _list.Items.Count - 1) return;
        MoveSelectedTo(delta < 0 ? indices[0] - 1 : indices[^1] + 2);
    }

    /// <summary>選んでいる行を、今の並びでの index の位置（その行の前）へまとめて動かす。</summary>
    private void MoveSelectedTo(int index)
    {
        var selected = _list.SelectedItems.Cast<ListViewItem>().OrderBy(i => i.Index).ToList();
        if (selected.Count == 0) return;

        // 動かしても並びが変わらないなら、何もしない（元に戻す履歴も増やさない）
        bool contiguous = selected[^1].Index - selected[0].Index == selected.Count - 1;
        if (contiguous && index >= selected[0].Index && index <= selected[^1].Index + 1) return;

        PushUndo();
        int before = selected.Count(i => i.Index < index);
        _list.BeginUpdate();
        foreach (var item in selected)
            _list.Items.Remove(item);
        index = Math.Clamp(index - before, 0, _list.Items.Count);
        for (int i = 0; i < selected.Count; i++)
        {
            _list.Items.Insert(index + i, selected[i]);
            selected[i].Selected = true;
        }
        _list.EndUpdate();

        selected[0].Focused = true;
        selected[0].EnsureVisible();
        RefreshAll();
    }

    private void SortByName()
    {
        if (_busy || _list.Items.Count < 2) return;
        PushUndo();
        var sorted = Items
            .OrderBy(i => EntryOf(i).FileName, NaturalStringComparer.Instance)
            .ThenBy(i => EntryOf(i).Path, NaturalStringComparer.Instance)
            .ToArray();
        _list.BeginUpdate();
        _list.Items.Clear();
        _list.Items.AddRange(sorted);
        _list.EndUpdate();
        RefreshAll();
        SetStatus("ファイル名の順に並べました。");
    }

    private ListState TakeSnapshot() =>
        new(Entries().Select(e => new Snapshot(e, e.Title, e.PageRange, e.Sheets, e.FitWidth, e.Rotations, e.OutlineDepth, e.OutlineChanges, e.PageBookmarks)).ToArray(),
            new Dictionary<string, string>(_folderNames, StringComparer.OrdinalIgnoreCase), _outlineDepth, _folderBookmarks);

    /// <summary>一覧を変える操作の直前に呼ぶ。</summary>
    private void PushUndo()
    {
        _undoStack.Add(TakeSnapshot());
        if (_undoStack.Count > UndoLimit) _undoStack.RemoveAt(0);
        _redoStack.Clear();
    }

    private void Undo() => Step(_undoStack, _redoStack, "元に戻しました。");

    private void Redo() => Step(_redoStack, _undoStack, "やり直しました。");

    private void Step(List<ListState> from, List<ListState> to, string message)
    {
        if (_busy || from.Count == 0) return;
        EndEdit(commit: false);
        to.Add(TakeSnapshot());
        var snapshot = from[^1];
        from.RemoveAt(from.Count - 1);
        Restore(snapshot);
        SetStatus(message);
    }

    private void Restore(ListState state)
    {
        _folderNames.Clear();
        foreach (var (key, name) in state.FolderNames) _folderNames[key] = name;
        _outlineDepth = state.OutlineDepth;
        _folderBookmarks = state.FolderBookmarks;
        var snapshot = state.Rows;
        var selected = SelectedEntries().ToHashSet();
        var kept = snapshot.Select(s => s.Entry).ToHashSet();
        foreach (var entry in Entries().Where(e => !kept.Contains(e)))
            Detach(entry);

        _list.BeginUpdate();
        _list.Items.Clear();
        foreach (var saved in snapshot)
        {
            var entry = saved.Entry;
            bool reconvert = entry.FitWidth != saved.FitWidth || !SameSheets(entry.Sheets, saved.Sheets);
            entry.Title = saved.Title;
            entry.PageRange = saved.PageRange;
            entry.Sheets = saved.Sheets;
            entry.FitWidth = saved.FitWidth;
            entry.Rotations = saved.Rotations;
            entry.OutlineDepth = saved.OutlineDepth;
            entry.OutlineChanges = saved.OutlineChanges;
            entry.PageBookmarks = saved.PageBookmarks;

            var item = CreateItem(entry);
            _list.Items.Add(item);
            item.Selected = selected.Contains(entry);
            if (reconvert) Prepare(entry);
            else EnsureFresh(entry);
        }
        _list.EndUpdate();
        RefreshAll();
    }

    private static bool SameSheets(IReadOnlyList<string>? a, IReadOnlyList<string>? b) =>
        a is null || b is null ? a is null && b is null : a.SequenceEqual(b);

    private void EditSelected(int column)
    {
        if (_list.SelectedItems.Count == 1) BeginEdit(_list.SelectedItems[0], column);
    }

    /// <summary>しおり名かページ指定のセルの上に入力欄を出して、その場で書き換えられるようにする。</summary>
    private void BeginEdit(ListViewItem item, int column)
    {
        if (_busy) return;
        EndEdit(commit: true);
        item.EnsureVisible();
        var entry = EntryOf(item);

        var bounds = item.SubItems[column].Bounds;
        bounds.Intersect(_list.ClientRectangle);
        if (bounds.Width < LogicalToDeviceUnits(40)) bounds.Width = LogicalToDeviceUnits(120);

        var editor = new TextBox
        {
            Text = column == ColTitle ? entry.Title : entry.PageRange,
            Tag = (entry, column),
            BorderStyle = BorderStyle.FixedSingle,
            PlaceholderText = column == ColRange ? "例: 1-3,5" : "",
        };
        editor.SetBounds(bounds.X, bounds.Y - 1, bounds.Width, bounds.Height + 2);
        editor.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (EndEdit(commit: true)) _list.Focus();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                EndEdit(commit: false);
                _list.Focus();
            }
        };
        // ほかの場所を押したときは、書きかけの内容で確定する（正しく読めない指定は捨てる）
        editor.LostFocus += (_, _) =>
        {
            if (ReferenceEquals(_editor, editor) && !EndEdit(commit: true)) EndEdit(commit: false);
        };

        _editor = editor;
        _list.Controls.Add(editor);
        editor.BringToFront();
        editor.Focus();
        editor.SelectAll();
        if (column == ColRange)
            SetStatus("使うページを「1-3,5」のように書きます。「8-」は 8 ページ以降。空にするとすべて。プレビューのページを押しても選べます。");
    }

    /// <returns>入力欄を閉じられたら true。書かれた内容が正しく読めず、開いたままにしたときは false。</returns>
    private bool EndEdit(bool commit)
    {
        if (_editor is not { } editor) return true;
        var (entry, column) = ((FileEntry, int))editor.Tag!;

        if (commit && column == ColRange && !PageRange.IsValid(editor.Text))
        {
            SystemSounds.Beep.Play();
            SetStatus("ページ指定は「1-3,5」のように、ページ番号をカンマとハイフンで書いてください。");
            return false;
        }

        _editor = null;
        string text = editor.Text;
        _list.Controls.Remove(editor);
        editor.Dispose();

        if (commit && ItemOf(entry) is not null)
        {
            if (column == ColTitle)
            {
                string title = text.Trim();
                if (title.Length == 0) title = entry.DefaultTitle;
                if (title != entry.Title)
                {
                    PushUndo();
                    entry.Title = title;
                    EntryChanged(entry);
                }
            }
            else
            {
                SetPageRange(entry, text);
            }
        }
        return true;
    }

    private void SetPageRange(FileEntry entry, string text)
    {
        if (_busy) return;
        string range = PageRange.Normalize(text);
        // 全ページを指す指定は、空（すべて）にそろえる
        if (entry.Converted is { } pdf && PageRange.Resolve(range, pdf.PageCount).Count == pdf.PageCount) range = "";
        if (range == entry.PageRange) return;
        PushUndo();
        entry.PageRange = range;
        EntryChanged(entry);

        if (entry.Converted is { } converted)
        {
            int used = entry.SelectedPages.Count;
            SetStatus(used == converted.PageCount ? $"{entry.FileName}: すべてのページを結合します。"
                : used == 0 ? $"{entry.FileName}: 指定したページがありません（全 {converted.PageCount} ページ）。"
                : $"{entry.FileName}: {converted.PageCount} ページ中 {used} ページを結合します。");
        }
    }

    /// <summary>プレビューのページが押されたとき、そのページを結合から外す（もう一度押すと戻す）。</summary>
    private void TogglePage(FileEntry entry, int pageIndex)
    {
        if (_busy || entry.Converted is not { } pdf) return;
        var pages = new SortedSet<int>(entry.SelectedPages);
        if (!pages.Remove(pageIndex)) pages.Add(pageIndex);
        if (pages.Count == 0)
        {
            SystemSounds.Beep.Play();
            SetStatus("最後の 1 ページは外せません。ファイルごと外すときは、リストから削除してください。");
            return;
        }
        SetPageRange(entry, PageRange.Format(pages, pdf.PageCount));
    }

    /// <summary>ページを回す。degrees は右へ 90、左へ -90。</summary>
    private void RotatePage(FileEntry entry, int pageIndex, int degrees)
    {
        if (_busy || entry.Converted is not { } pdf || pageIndex < 0 || pageIndex >= pdf.PageCount) return;
        var rotations = new Dictionary<int, int>(entry.Rotations);
        int turn = PageGlyphs.Normalize(rotations.GetValueOrDefault(pageIndex) + degrees);
        if (turn == 0) rotations.Remove(pageIndex);
        else rotations[pageIndex] = turn;

        PushUndo();
        entry.Rotations = rotations;
        EntryChanged(entry);
        SetStatus($"{entry.FileName}: {pageIndex + 1} ページを{(degrees > 0 ? "右" : "左")}に回しました" + (turn == 0 ? "（元の向き）。" : $"（{turn}°）。"));
    }

    /// <summary>ページを大きく表示する。作成中は見るだけにする。</summary>
    private void ShowPage(FileEntry entry, int pageIndex)
    {
        if (entry.Converted is null) return;
        using var viewer = new PageViewer(entry, pageIndex, editable: !_busy,
            toggle: page => TogglePage(entry, page),
            rotate: (page, degrees) => RotatePage(entry, page, degrees),
            addBookmark: (page, owner) => AddBookmark(entry, page, owner));
        viewer.ShowDialog(this);
    }

    /// <summary>
    /// ページを見て、そのページにしおりを足す（元のしおりのない、スキャンした PDF などのため）。
    /// 名前を尋ね、元に戻す（Ctrl+Z）で戻せるようにする。名前の変更や削除は「しおりを確認」で行う。
    /// </summary>
    private void AddBookmark(FileEntry entry, int page, IWin32Window? owner = null)
    {
        if (_busy || entry.Converted is not { } pdf || page < 0 || page >= pdf.PageCount) return;
        using var dialog = new TextPromptDialog("しおりを付ける", $"{entry.FileName} の {page + 1} ページに付けるしおりの名前", "");
        if (dialog.ShowDialog(owner ?? this) != DialogResult.OK || dialog.Value.Length == 0) return;

        PushUndo();
        // 同じページに足したしおりは、足した順に並べる（OrderBy は順を保つ）
        var items = entry.AddedBookmarks.Append(new AddedBookmark(dialog.Value, page)).OrderBy(b => b.Page).ToList();
        entry.PageBookmarks = new PageBookmarks(pdf.PageSizes, items);
        EntryChanged(entry);
        SetStatus(entry.SelectedPages.Contains(page)
            ? $"{page + 1} ページにしおり「{dialog.Value}」を付けました。名前の変更や削除は「しおりを確認」でできます。"
            : $"{page + 1} ページにしおり「{dialog.Value}」を付けました。ただし、このページは結合しないので、結合した PDF には付きません。");
    }

    private void ExtractSelected()
    {
        if (SelectedEntries().Take(2).ToList() is [var entry]) ExtractPages(entry);
    }

    /// <summary>
    /// 1 ファイルから選んだページだけを 1 つの PDF に保存する。結合の一覧そのものは変えない。
    /// pages を渡すと、そのページを選んだ状態で開く（起動の指定 --pick から）。
    /// </summary>
    private void ExtractPages(FileEntry entry, string pages = "")
    {
        if (_busy || entry.Converted is null) return;
        using var dialog = new PageSelectDialog(entry);
        if (pages.Length > 0) dialog.Preselect(pages);
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Written is not { } written) return;
        SetCreatedStatus(written, 1, dialog.WrittenPages);
        if (_optOpen.Checked) ShellOpen(written);
    }

    /// <summary>一覧でその行だけを選び、見える位置まで送る（プレビューにもその行が出る）。</summary>
    private void SelectEntry(FileEntry entry)
    {
        foreach (ListViewItem item in _list.Items)
        {
            bool match = ReferenceEquals(item.Tag, entry);
            item.Selected = match;
            if (!match) continue;
            item.Focused = true;
            item.EnsureVisible();
        }
    }

    private void SplitSelected()
    {
        if (SelectedEntries().Take(2).ToList() is [var entry]) Split(entry);
    }

    /// <summary>1 ファイルを複数の PDF に分ける。結合の一覧そのものは変えない。</summary>
    private void Split(FileEntry entry)
    {
        if (_busy || !entry.IsUsable) return;
        using var dialog = new SplitDialog(entry);
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Written.Count == 0) return;
        SetFilesStatus(dialog.Written, $"{dialog.Written.Count} ファイルに分割しました");
        if (dialog.RevealAfterwards) Reveal(dialog.Written[0]);
    }

    private void ChooseSheetsForSelected()
    {
        if (SelectedEntries().FirstOrDefault() is { } entry) ChooseSheets(entry);
    }

    private void ChooseSheets(FileEntry entry)
    {
        if (_busy || entry.Converted?.Sheets is not { } sheets) return;
        using var dialog = new ExcelSheetsDialog(entry.FileName, sheets, entry.FitWidth);
        if (dialog.ShowDialog(this) != DialogResult.OK) return;

        var chosen = dialog.SelectedSheets;
        if (SameSheets(chosen, entry.Sheets) && dialog.FitWidth == entry.FitWidth) return;
        PushUndo();
        entry.Sheets = chosen;
        entry.FitWidth = dialog.FitWidth;
        // シートや幅が変わるとページの数も変わるので、前のページ指定と回転は引き継がない
        entry.PageRange = "";
        entry.Rotations = FileEntry.NoRotations;
        Prepare(entry);
    }

    private void AskPasswordForSelected()
    {
        if (SelectedEntries().FirstOrDefault() is { } entry) AskPassword(entry);
    }

    private void AskPassword(FileEntry entry)
    {
        if (_busy) return;
        using var dialog = new PasswordDialog(entry.FileName, entry.Password is null ? null : "パスワードが違います。もう一度入力してください。");
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Password.Length == 0) return;
        entry.Password = dialog.Password;
        Prepare(entry);
    }

    private void OnSelectionChanged()
    {
        UpdateButtons();
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        var selected = SelectedEntries().Take(2).ToList();
        _preview.Display(selected.Count == 1 ? selected[0] : null, _list.SelectedItems.Count, _list.Items.Count, locked: _busy);
    }

    private void OpenSelected()
    {
        foreach (var entry in SelectedEntries().Take(5).ToList())
            ShellOpen(entry.Path);
    }

    private void RevealSelected()
    {
        if (SelectedEntries().FirstOrDefault() is { } entry) Reveal(entry.Path);
    }

    private void OpenConvertedSelected()
    {
        if (SelectedEntries().FirstOrDefault()?.Converted is { } pdf) ShellOpen(pdf.PdfPath);
    }

    private void OnListDoubleClick(object? sender, MouseEventArgs e)
    {
        var hit = _list.HitTest(e.Location);
        if (hit.Item is null) return;
        int column = hit.SubItem is null ? ColName : hit.Item.SubItems.IndexOf(hit.SubItem);
        var entry = EntryOf(hit.Item);

        if (column is ColTitle or ColRange) BeginEdit(hit.Item, column);
        else if (entry.State == EntryState.NeedsPassword) AskPassword(entry);
        else ShellOpen(entry.Path);
    }

    private void OnListKeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Delete:
                RemoveSelected();
                break;
            case Keys.Up when e.Alt:
                MoveSelectedBy(-1);
                break;
            case Keys.Down when e.Alt:
                MoveSelectedBy(+1);
                break;
            case Keys.A when e.Control:
                _list.BeginUpdate();
                foreach (var item in Items) item.Selected = true;
                _list.EndUpdate();
                break;
            case Keys.F2:
                EditSelected(ColTitle);
                break;
            case Keys.F3:
                EditSelected(ColRange);
                break;
            default:
                return;
        }
        e.Handled = true;
        e.SuppressKeyPress = true;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        // 入力欄で書いている間は、貼り付けや元に戻すを入力欄に任せる
        if (_editor is null)
        {
            switch (keyData)
            {
                case Keys.Control | Keys.O:
                    ChooseFiles();
                    return true;
                case Keys.Control | Keys.V:
                    PasteFiles();
                    return true;
                case Keys.Control | Keys.Z:
                    Undo();
                    return true;
                case Keys.Control | Keys.Y:
                case Keys.Control | Keys.Shift | Keys.Z:
                    Redo();
                    return true;
                case Keys.Control | Keys.Enter:
                    if (_run.Enabled) _run.PerformClick();
                    return true;
                case Keys.Escape when IsScanning:
                    CancelScan();
                    return true;
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void OnItemDrag(object? sender, ItemDragEventArgs e)
    {
        if (_busy || _list.SelectedItems.Count == 0) return;
        EndEdit(commit: true);
        _list.DoDragDrop(new DataObject(RowDragFormat, true), DragDropEffects.Move);
        HideInsertMark();
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        bool rows = e.Data?.GetDataPresent(RowDragFormat) == true;
        bool files = e.Data?.GetDataPresent(DataFormats.FileDrop) == true;
        if (_busy || (!rows && !files))
        {
            e.Effect = DragDropEffects.None;
            HideInsertMark();
            return;
        }

        e.Effect = rows ? DragDropEffects.Move : DragDropEffects.Copy;
        AutoScrollWhileDragging(e.Y);
        ShowInsertMark(InsertIndexAt(e.X, e.Y));
    }

    private void OnDragDrop(object? sender, DragEventArgs e)
    {
        HideInsertMark();
        if (_busy) return;
        int index = InsertIndexAt(e.X, e.Y);

        if (e.Data?.GetDataPresent(RowDragFormat) == true)
        {
            MoveSelectedTo(index);
        }
        else if (e.Data?.GetData(DataFormats.FileDrop) is string[] paths)
        {
            // ドロップ元（エクスプローラー）を待たせないよう、処理は後で行う
            BeginInvoke(() =>
            {
                Activate();
                _ = AddPathsAsync(paths, index);
            });
        }
    }

    /// <summary>画面座標の位置に落としたとき、何行目の前に入るか。</summary>
    private int InsertIndexAt(int screenX, int screenY)
    {
        int count = _list.Items.Count;
        if (count == 0) return 0;

        var p = _list.PointToClient(new Point(screenX, screenY));
        var item = _list.GetItemAt(LogicalToDeviceUnits(4), p.Y);
        if (item is not null)
        {
            var b = item.Bounds;
            return p.Y < b.Top + b.Height / 2 ? item.Index : item.Index + 1;
        }

        var top = _list.TopItem ?? _list.Items[0];
        return p.Y < top.Bounds.Top ? top.Index : count;
    }

    private void ShowInsertMark(int index)
    {
        int count = _list.Items.Count;
        if (count == 0)
        {
            HideInsertMark();
            return;
        }
        int y = index < count ? _list.Items[index].Bounds.Top : _list.Items[count - 1].Bounds.Bottom;
        int height = LogicalToDeviceUnits(2);
        _insertMark.SetBounds(0, y - height / 2, _list.ClientSize.Width, height);
        _insertMark.Visible = true;
        _insertMark.BringToFront();
    }

    private void HideInsertMark() => _insertMark.Visible = false;

    private void AutoScrollWhileDragging(int screenY)
    {
        var top = _list.TopItem;
        if (top is null) return;
        int y = _list.PointToClient(new Point(0, screenY)).Y;
        int rowHeight = Math.Max(1, top.Bounds.Height);

        if (y < top.Bounds.Top + rowHeight && top.Index > 0)
        {
            _list.EnsureVisible(top.Index - 1);
        }
        else if (y > _list.ClientSize.Height - rowHeight)
        {
            int visible = Math.Max(1, (_list.ClientSize.Height - top.Bounds.Top) / rowHeight);
            _list.EnsureVisible(Math.Min(_list.Items.Count - 1, top.Index + visible));
        }
    }

    /// <summary>1 行の状態が変わったときに呼ぶ。その行と、合計・ボタン・プレビューを更新する。</summary>
    private void EntryChanged(FileEntry entry)
    {
        if (ItemOf(entry) is { } item) UpdateItem(item);
        UpdateSummary();
        UpdateButtons();
        if (ReferenceEquals(_preview.Entry, entry)
            || (_list.SelectedItems.Count == 1 && ReferenceEquals(EntryOf(_list.SelectedItems[0]), entry)))
        {
            UpdatePreview();
        }
    }

    private void RefreshAll()
    {
        _list.BeginUpdate();
        foreach (var item in Items) UpdateItem(item);
        _list.EndUpdate();
        _emptyHint.Visible = _list.Items.Count == 0;
        FitColumns(); // 行が増えて縦のスクロールバーが出入りすると、使える幅が変わる
        UpdateSummary();
        UpdateButtons();
        UpdatePreview();
    }

    private static void UpdateItem(ListViewItem item)
    {
        var entry = EntryOf(item);
        int used = entry.SelectedPages.Count;
        int total = entry.Converted?.PageCount ?? 0;

        SetText(item, ColName, entry.FileName);
        SetText(item, ColTitle, entry.Title);
        SetText(item, ColPages, entry.Converted is null ? "" : used == total ? $"{total}" : $"{used}/{total}");
        SetText(item, ColRange, entry.PageRange);
        SetText(item, ColState, StateText(entry));
        // 上から順に結合する。並べ替えや削除のあとは RefreshAll で付け直す
        SetText(item, ColOrder, item.Index >= 0 ? $"{item.Index + 1}" : "");

        bool problem = entry.State is EntryState.Failed or EntryState.NeedsPassword
            || (entry.State == EntryState.Ready && used == 0);
        var color = problem ? Theme.Danger : entry.IsBusy ? SystemColors.GrayText : SystemColors.WindowText;
        if (item.ForeColor != color) item.ForeColor = color;

        // 行に乗せたときに、場所と、結合できない理由の全文を出す（列の幅では切れてしまうため）
        string tip = problem && entry.State != EntryState.Ready ? $"{entry.Path}\r\n{entry.Error}" : entry.Path;
        if (item.ToolTipText != tip) item.ToolTipText = tip;
    }

    private static string StateText(FileEntry entry)
    {
        switch (entry.State)
        {
            case EntryState.Queued:
                return "待機中";
            case EntryState.Converting:
                return "変換中...";
            case EntryState.NeedsPassword:
                return entry.Password is null ? "パスワードが必要" : "パスワードが違います";
            case EntryState.Failed:
                return $"失敗: {entry.Error}";
        }

        if (entry.SelectedPages.Count == 0) return "指定したページがありません";
        var notes = new List<string>();
        if (entry.Converted?.Sheets is { } sheets && sheets.Any(s => !s.Included))
            notes.Add($"シート {sheets.Count(s => s.Included)}/{sheets.Count}");
        if (entry.FitWidth) notes.Add("幅に収める");
        if (entry.RotatedCount > 0) notes.Add($"{entry.RotatedCount} ページ回転");
        return notes.Count > 0 ? string.Join("・", notes) : "OK";
    }

    private static void SetText(ListViewItem item, int column, string text)
    {
        if (item.SubItems[column].Text != text) item.SubItems[column].Text = text;
    }
}
