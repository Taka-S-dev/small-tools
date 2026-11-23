using DocsToPdf.Core;

namespace DocsToPdf.Ui;

// 「一覧を保存」「一覧を開く」: 並び・行ごとの決めごと・オプションをファイルに残し、あとで同じ PDF を作り直す
internal sealed partial class MainForm
{
    private const string CompositionFilter = "DocsToPdf の一覧|*" + Composition.Extension;

    /// <summary>最後に開いた・保存した一覧のファイル。次に保存するときの既定にする。</summary>
    private string? _compositionPath;

    private void SaveComposition()
    {
        if (_busy || _list.Items.Count == 0) return;
        EndEdit(commit: true);
        var entries = Entries().ToList();

        using var dialog = new SaveFileDialog
        {
            Filter = CompositionFilter,
            Title = "一覧を保存",
            DefaultExt = Composition.Extension.TrimStart('.'),
            AddExtension = true,
            OverwritePrompt = true,
        };
        if (_compositionPath is { } last)
        {
            dialog.InitialDirectory = Path.GetDirectoryName(last);
            dialog.FileName = Path.GetFileName(last);
        }
        else
        {
            dialog.InitialDirectory = Path.GetDirectoryName(entries[0].Path);
            dialog.FileName = Path.GetFileNameWithoutExtension(entries[0].Path) + Composition.Extension;
        }
        if (dialog.ShowDialog(this) != DialogResult.OK) return;

        var composition = new Composition(entries.Select(CompositionItem.Of).ToList(),
            _optPageNumbers.Checked, _optToc.Checked, _optDuplex.Checked, _outlineDepth, _folderBookmarks,
            new Dictionary<string, string>(_folderNames, StringComparer.OrdinalIgnoreCase))
        {
            Dividers = _optDividers.Checked,
        };
        try
        {
            CompositionFile.Save(dialog.FileName, composition);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            ShowError("一覧を保存できませんでした", ex.Message);
            return;
        }
        _compositionPath = dialog.FileName;
        bool passwords = entries.Any(e => e.Password is not null);
        SetStatus($"一覧を保存しました: {Path.GetFileName(dialog.FileName)}"
            + (passwords ? "（パスワードは保存していません。開いたときにもう一度入力してください）" : ""));
    }

    private void ChooseComposition()
    {
        if (_busy) return;
        using var dialog = new OpenFileDialog { Filter = CompositionFilter, Title = "一覧を開く" };
        if (_compositionPath is { } last) dialog.InitialDirectory = Path.GetDirectoryName(last);
        if (dialog.ShowDialog(this) == DialogResult.OK)
            OpenComposition(dialog.FileName);
    }

    /// <summary>
    /// 一覧が、ほかのコンピューターの共有フォルダを指しているときに、そこへ探しに行ってよいかを尋ねる。
    /// スクリプトから使ったとき（--out）は尋ねられないので、使わない。
    /// </summary>
    private bool ConfirmNetworkPaths(IReadOnlyList<string> hosts)
    {
        if (_startup.IsBatch) return false;
        var open = new TaskDialogButton("開く");
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Icon = TaskDialogIcon.Warning,
            Heading = "ほかのコンピューターのファイルを使う一覧です",
            Text = $"この一覧は、次の共有フォルダにあるファイルを指しています。\n{string.Join("\n", hosts.Select(h => @"\\" + h))}\n\n"
                + "開くと、そのコンピューターへ Windows のサインイン情報が送られます。人から受け取った一覧なら、心当たりのある場所かを確かめてください。",
            Buttons = { open, TaskDialogButton.Cancel },
            DefaultButton = TaskDialogButton.Cancel,
        };
        return TaskDialog.ShowDialog(this, page) == open;
    }

    /// <summary>保存した一覧で、今の一覧を置き換える（Ctrl+Z で前の一覧に戻せる）。見つからないファイルは除いて知らせる。</summary>
    private void OpenComposition(string path)
    {
        if (_busy)
        {
            SetStatus("作成中のため、一覧を開けませんでした。終わってからもう一度開いてください。");
            return;
        }
        Composition composition;
        try
        {
            composition = CompositionFile.Load(path, ConfirmNetworkPaths);
        }
        catch (NetworkPathsException ex)
        {
            if (_startup.IsBatch) ShowError("一覧を開けませんでした", ex.Message + "\nDocsToPdfCli merge に --allow-network を付けると使えます。");
            else SetStatus("一覧を開くのをやめました。");
            return;
        }
        catch (Exception ex) when (ex is FormatException or IOException or UnauthorizedAccessException)
        {
            ShowError("一覧を開けませんでした", ex.Message);
            return;
        }

        var missing = composition.Items.Where(i => !File.Exists(i.Path) || !FileEntry.TryGetKind(i.Path, out _)).Select(i => i.Path).ToList();
        if (_startup.IsBatch && missing.Count > 0)
        {
            // スクリプトから使ったときは、足りないまま作らずに失敗にする
            ShowError($"{missing.Count} 件のファイルが見つかりません", string.Join("\n", missing));
            return;
        }

        var entries = new List<FileEntry>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in composition.Items)
        {
            if (!File.Exists(item.Path) || !FileEntry.TryGetKind(item.Path, out var kind)) continue;
            if (!seen.Add(item.Path)) continue;
            entries.Add(new FileEntry(item.Path, kind)
            {
                Root = item.Root,
                Title = item.Title,
                PageRange = item.PageRange,
                Sheets = item.Sheets,
                FitWidth = item.FitWidth,
                Rotations = item.Rotations,
                OutlineDepth = item.OutlineDepth,
                OutlineChanges = item.OutlineChanges,
                PageBookmarks = item.PageBookmarks,
            });
        }

        EndEdit(commit: true);
        PushUndo();
        foreach (var entry in Entries()) Detach(entry);
        _list.BeginUpdate();
        _list.Items.Clear();
        foreach (var entry in entries) _list.Items.Add(CreateItem(entry));
        _list.EndUpdate();

        // スクリプトから使ったとき（--out）は、起動の指定で足したオプションも生かす
        _optPageNumbers.Checked = composition.PageNumbers || (_startup.IsBatch && _startup.PageNumbers);
        _optToc.Checked = composition.TableOfContents || (_startup.IsBatch && _startup.TableOfContents);
        _optDuplex.Checked = composition.DuplexBlanks || (_startup.IsBatch && _startup.DuplexBlanks);
        _optDividers.Checked = composition.Dividers || (_startup.IsBatch && _startup.Dividers);
        _outlineDepth = composition.OutlineDepth;
        _folderBookmarks = composition.FolderBookmarks;
        _folderNames.Clear();
        foreach (var (key, folderName) in composition.FolderNames) _folderNames[key] = folderName;
        _compositionPath = Path.GetFullPath(path);

        foreach (var entry in entries) Prepare(entry);
        RefreshAll();

        string name = Path.GetFileName(path);
        if (missing.Count > 0)
        {
            SetStatus($"一覧を開きました: {name}（見つからない {missing.Count} 件は除きました。Ctrl+Z で前の一覧に戻せます）");
            ShowError($"{missing.Count} 件のファイルが見つかりません",
                "移動したか、名前を変えた可能性があります。見つかったファイルだけで一覧を開きました。\n\n"
                + string.Join("\n", missing.Take(20)) + (missing.Count > 20 ? $"\n…ほか {missing.Count - 20} 件" : ""));
        }
        else
        {
            SetStatus($"一覧を開きました: {name}（{entries.Count} 件。Ctrl+Z で前の一覧に戻せます）");
        }
    }
}
