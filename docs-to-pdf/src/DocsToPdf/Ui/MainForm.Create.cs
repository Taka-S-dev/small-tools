using DocsToPdf.Core;

namespace DocsToPdf.Ui;

// 「結合して PDF を作成」: 残っている変換を待ち、結合して保存する
// 「個別に PDF を保存」: 残っている変換を待ち、ファイルごとに PDF を保存する
internal sealed partial class MainForm
{
    /// <summary>
    /// 一覧の順に結合して保存する。
    /// </summary>
    /// <param name="outputPath">保存先。null なら尋ねる。指定があるのは、スクリプトから使ったとき（--out）。</param>
    /// <returns>PDF を作れたら true。</returns>
    private Task<bool> CreateAsync(string? outputPath) => RunExclusiveAsync(async ct =>
    {
        bool batch = outputPath is not null;
        var entries = Entries().ToList();
        string? output = outputPath ?? AskOutputPath(entries);
        if (output is null) return false;
        ShowProgress(true);

        await WaitForConversionsAsync(entries, ct);

        var usable = entries.Where(e => e.IsUsable).ToList();
        var unusable = entries.Where(e => !e.IsUsable).ToList();
        if (usable.Count == 0)
        {
            SetStatus("結合できるファイルがありませんでした。");
            ShowError("どのファイルも結合できませんでした", ProblemList(unusable));
            return false;
        }
        if (unusable.Count > 0)
        {
            if (batch)
            {
                ShowError($"{unusable.Count} 件のファイルを結合できません", ProblemList(unusable));
                return false;
            }
            if (!ConfirmPartial(unusable, $"{unusable.Count} 件のファイルを結合できません",
                    "結合できるファイルだけで PDF を作りますか？", $"残りの {usable.Count} 件で作成"))
            {
                SetStatus("作成を取りやめました。");
                return false;
            }
        }

        SetStatus("PDF を結合しています...");
        var parts = PartsOf(usable, CurrentChoices(usable));
        var options = CurrentOptions(Path.GetFileNameWithoutExtension(output));
        int pages = await Task.Run(() => PdfMerger.Merge(parts, options, output, StartProgress(parts.Count), ct), ct);

        _settings.LastOutputDirectory = Path.GetDirectoryName(output);
        SetCreatedStatus(output, usable.Count, pages);
        WindowAttention.Flash(this);
        if (_optOpen.Checked && !_closeWhenStopped) ShellOpen(output);
        return true;
    });

    /// <summary>
    /// 結合せずに、ファイルごとに「元のファイル名.pdf」を保存する。
    /// プレビューで外したページは含めず、回したページは回した向きで保存する。ページ番号はオプションに従い、目次と白紙は付けない。
    /// </summary>
    private Task<bool> SaveEachAsync() => RunExclusiveAsync(async ct =>
    {
        var entries = Entries().ToList();
        if (!AskSaveEachTarget(entries, out string? folder)) return false;
        ShowProgress(true);

        await WaitForConversionsAsync(entries, ct);

        var usable = entries.Where(e => e.IsUsable).ToList();
        var unusable = entries.Where(e => !e.IsUsable).ToList();
        if (usable.Count == 0)
        {
            SetStatus("PDF にできるファイルがありませんでした。");
            ShowError("どのファイルも PDF にできませんでした", ProblemList(unusable));
            return false;
        }
        if (unusable.Count > 0 && !ConfirmPartial(unusable, $"{unusable.Count} 件のファイルを PDF にできません",
                "PDF にできるファイルだけを保存しますか？", $"残りの {usable.Count} 件を保存"))
        {
            SetStatus("保存を取りやめました。");
            return false;
        }

        // 一覧のどのファイル（PDF にできなかったものや、元の PDF を含む）にも書かない
        var names = IndividualPdf.FileNames(usable.Select(e => e.Path).ToList(),
            source => folder ?? Path.GetDirectoryName(source)!, entries.Select(e => e.Path));
        var existing = names.Where(File.Exists).ToList();
        if (existing.Count > 0 && !ConfirmOverwrite(existing))
        {
            SetStatus("保存を取りやめました。");
            return false;
        }

        SetStatus("PDF を保存しています...");
        var parts = usable.Select(PartOf).ToList();
        bool pageNumbers = _optPageNumbers.Checked;
        int? depth = _outlineDepth;
        var progress = StartProgress(parts.Count);
        await Task.Run(() =>
        {
            for (int i = 0; i < parts.Count; i++)
            {
                IndividualPdf.Write(parts[i], names[i], pageNumbers, ct, depth);
                ((IProgress<int>)progress).Report(i + 1);
            }
        }, ct);

        if (folder is not null) _settings.LastOutputDirectory = folder;
        SetFilesStatus(names, folder is null ? $"{names.Count} ファイルを、それぞれ元のファイルと同じフォルダに保存しました" : $"{names.Count} ファイルを保存しました");
        WindowAttention.Flash(this);
        if (_optOpen.Checked && !_closeWhenStopped) Reveal(names[0]);
        return true;
    });

    /// <summary>今の一覧と設定でのしおりの選び方（ファイルの番号は entries の並び）。</summary>
    private BookmarkChoices CurrentChoices(IReadOnlyList<FileEntry> entries)
    {
        var choices = new BookmarkChoices
        {
            Depth = _outlineDepth,
            Folders = _folderBookmarks,
            FolderNames = new(_folderNames, StringComparer.OrdinalIgnoreCase),
        };
        for (int i = 0; i < entries.Count; i++)
        {
            choices.Titles[i] = entries[i].Title;
            choices.FileDepths[i] = entries[i].OutlineDepth;
            choices.OutlineEdits[i] = new(entries[i].OutlineEdits);
            choices.Added[i] = [.. entries[i].AddedBookmarks];
        }
        return choices;
    }

    /// <summary>結合する行を、しおりの選び方（名前・段数・フォルダの段）を当てて並べる。</summary>
    private List<MergePart> PartsOf(IReadOnlyList<FileEntry> entries, BookmarkChoices choices) =>
        ApplyChoices(entries, entries.Select(PartOf).ToList(), Entries().Select(e => e.Root).ToList(), choices);

    /// <param name="parts">entries を PartOf にしたもの（しおりの確認では、開いたときの写しを使い回す）。</param>
    /// <param name="allRoots">一覧全体の行のフォルダ（入れたフォルダが 1 つだけかを、変換中の行も含めて決める）。</param>
    private static List<MergePart> ApplyChoices(IReadOnlyList<FileEntry> entries, IReadOnlyList<MergePart> parts,
        IReadOnlyList<string?> allRoots, BookmarkChoices choices)
    {
        var groups = choices.Folders
            ? BookmarkFolders.Of(entries.Select(e => e.Path).ToList(), entries.Select(e => e.Root).ToList(), allRoots)
            : null;
        return entries.Select((e, i) => parts[i] with
        {
            Title = TitleOf(e, choices.Titles.GetValueOrDefault(i, e.Title)),
            OutlineDepth = choices.FileDepths.GetValueOrDefault(i, e.OutlineDepth),
            OutlineEdits = choices.OutlineEdits.TryGetValue(i, out var edits) ? edits : parts[i].OutlineEdits,
            Added = choices.Added.TryGetValue(i, out var added) ? added : parts[i].Added,
            Folders = groups?[i].Select(f => choices.FolderNames.TryGetValue(f.Key, out var name) ? f with { Name = name } : f).ToList() ?? [],
        }).ToList();
    }

    /// <summary>空の名前は、元のファイル名に戻す（一覧で名前を消したときと同じ）。</summary>
    private static string TitleOf(FileEntry e, string title) => title.Trim() is { Length: > 0 } t ? t : e.DefaultTitle;

    private static MergePart PartOf(FileEntry e) =>
        new(e.Converted!.PdfPath, e.Title, e.SelectedPages, e.Converted.Outline, e.Rotations) { OutlineDepth = e.OutlineDepth, OutlineEdits = e.OutlineEdits, Added = e.AddedBookmarks };

    /// <summary>進み具合のバーを 0 から始め、1 件終わるごとに進める報告先を返す。</summary>
    private Progress<int> StartProgress(int total)
    {
        _progress.Maximum = total;
        _progress.Value = 0;
        return new Progress<int>(done =>
        {
            _progress.Value = Math.Min(_progress.Maximum, done);
            _taskbar.Set(Handle, done, total);
        });
    }

    /// <summary>
    /// 一覧を変えられない状態にして work を行い、中止と保存の失敗を画面に出す。
    /// 保存先を選んでいる間も含めて一覧を止める（選んだ後で行が増減しないように）。元のファイルが変わっていたら、先に変換し直す。
    /// </summary>
    private async Task<bool> RunExclusiveAsync(Func<CancellationToken, Task<bool>> work)
    {
        if (_busy || _list.Items.Count == 0) return false;
        EndEdit(commit: true);

        var cts = _cts = new CancellationTokenSource();
        SetBusy(true);
        ShowProgress(false);
        try
        {
            await _sourceCheck;
            await CheckSourcesAsync();
            return await work(cts.Token);
        }
        catch (OperationCanceledException)
        {
            SetStatus("中止しました。");
            return false;
        }
        catch (IOException ex)
        {
            SetStatus("保存できませんでした。");
            ShowError("PDF を保存できませんでした",
                $"{ex.Message}\r\n\r\n保存先の PDF を別のアプリで開いている場合は、閉じてからやり直してください。");
            return false;
        }
        catch (UnauthorizedAccessException ex)
        {
            SetStatus("保存できませんでした。");
            ShowError("PDF を保存できませんでした", $"{ex.Message}\r\n\r\n別の保存先を指定してください。");
            return false;
        }
        catch (Exception ex)
        {
            SetStatus("エラーで止まりました。");
            ShowError("処理中にエラーが起きました", ex.Message);
            return false;
        }
        finally
        {
            _cts = null;
            cts.Dispose();
            SetBusy(false);
            if (_closeWhenStopped) BeginInvoke(Close);
        }
    }

    /// <summary>すべての行の変換が終わる（成功か失敗が決まる）まで待つ。進み具合を表示する。</summary>
    private async Task WaitForConversionsAsync(IReadOnlyList<FileEntry> entries, CancellationToken ct)
    {
        _progress.Maximum = entries.Count;
        while (true)
        {
            int done = entries.Count(e => !e.IsBusy);
            _progress.Value = done;
            _taskbar.Set(Handle, done, entries.Count);
            if (done == entries.Count) return;

            var current = entries.FirstOrDefault(e => e.State == EntryState.Converting) ?? entries.First(e => e.IsBusy);
            SetStatus($"PDF に変換しています ({done + 1}/{entries.Count}): {current.FileName}");
            await Task.Delay(100, ct);
        }
    }

    private string? AskOutputPath(IReadOnlyList<FileEntry> entries)
    {
#if DEBUG
        // 画面の自動操作で確かめるとき用（Debug ビルドだけ）。保存先を尋ねるダイアログは自動操作しにくいので、環境変数で渡せる
        if (Environment.GetEnvironmentVariable("DOCSTOPDF_TEST_SAVE_AS") is { Length: > 0 } forced)
            return Path.GetFullPath(forced);
#endif

        string firstDir = Path.GetDirectoryName(entries[0].Path)!;
        string initialDir = Directory.Exists(_settings.LastOutputDirectory) ? _settings.LastOutputDirectory! : firstDir;

        using var dialog = new SaveFileDialog
        {
            Title = "結合した PDF の保存先",
            Filter = "PDF ファイル|*.pdf",
            DefaultExt = "pdf",
            AddExtension = true,
            OverwritePrompt = true,
            InitialDirectory = initialDir,
            FileName = $"{entries[0].DefaultTitle}_結合.pdf",
        };

        while (dialog.ShowDialog(this) == DialogResult.OK)
        {
            string path = Path.GetFullPath(dialog.FileName);
            if (!entries.Any(e => string.Equals(e.Path, path, StringComparison.OrdinalIgnoreCase)))
                return path;
            MessageBox.Show(this, "結合するファイルと同じ名前では保存できません。別の名前を指定してください。",
                AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        return null;
    }

    private void CancelCreate()
    {
        if (IsScanning)
        {
            CancelScan();
            return;
        }
        if (_cts is null) return;
        _cts.Cancel();
        _cancel.Enabled = false;
        SetStatus("中止しています...");
    }

    /// <summary>作成中は一覧やオプションを変えられないようにし、合計の代わりに進み具合を出す。</summary>
    private void SetBusy(bool busy)
    {
        _busy = busy;
        ShowProgress(busy);
        UpdateButtons();
        UpdatePreview();
    }

    private void ShowProgress(bool show)
    {
        _progress.Visible = show;
        _summary.Visible = !show;
        if (show) _taskbar.SetBusy(Handle);
        else _taskbar.Clear(Handle);
    }

    private bool ConfirmPartial(IReadOnlyList<FileEntry> unusable, string heading, string text, string proceed)
    {
        var go = new TaskDialogButton(proceed);
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = heading,
            Text = text,
            Icon = TaskDialogIcon.Warning,
            Expander = new TaskDialogExpander(ProblemList(unusable)) { Expanded = true, Position = TaskDialogExpanderPosition.AfterText },
            Buttons = { go, TaskDialogButton.Cancel },
            DefaultButton = go,
        };
        return TaskDialog.ShowDialog(this, page) == go;
    }

    /// <summary>保存先にすでに同じ名前の PDF があるとき、上書きしてよいかを尋ねる。既定は「キャンセル」。</summary>
    private bool ConfirmOverwrite(IReadOnlyList<string> existing)
    {
        var overwrite = new TaskDialogButton("上書きして保存");
        string list = string.Join("\r\n", existing.Take(15).Select(p => "・" + Path.GetFileName(p)))
            + (existing.Count > 15 ? $"\r\nほか {existing.Count - 15} 件" : "");
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = $"{existing.Count} 件の PDF が保存先にすでにあります",
            Text = "上書きしますか？",
            Icon = TaskDialogIcon.Warning,
            Expander = new TaskDialogExpander(list) { Expanded = true, Position = TaskDialogExpanderPosition.AfterText },
            Buttons = { overwrite, TaskDialogButton.Cancel },
            DefaultButton = TaskDialogButton.Cancel,
        };
        return TaskDialog.ShowDialog(this, page) == overwrite;
    }

    /// <summary>
    /// ファイルごとの PDF をどこに保存するかを尋ねる。「元のファイルと同じフォルダ」なら folder は null。
    /// 前回選んだほうを既定にする（Enter だけで同じ保存先をくり返せるように）。
    /// </summary>
    /// <returns>保存先が決まったら true。キャンセルなら false。</returns>
    private bool AskSaveEachTarget(IReadOnlyList<FileEntry> entries, out string? folder)
    {
        folder = null;
#if DEBUG
        // 画面の自動操作で確かめるとき用（Debug ビルドだけ）。"beside" なら元のファイルと同じフォルダ
        if (Environment.GetEnvironmentVariable("DOCSTOPDF_TEST_SAVE_FOLDER") is { Length: > 0 } forced)
        {
            if (forced != "beside") folder = Path.GetFullPath(forced);
            return true;
        }
#endif

        var beside = new TaskDialogCommandLinkButton("元のファイルと同じフォルダ", "それぞれのファイルの隣に「元のファイル名.pdf」を保存します");
        var choose = new TaskDialogCommandLinkButton("フォルダを選ぶ...", "すべてを 1 つのフォルダにまとめて保存します");
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = $"{entries.Count} ファイルの PDF をどこに保存しますか？",
            Buttons = { beside, choose, TaskDialogButton.Cancel },
            DefaultButton = _settings.SaveEachBesideSource ? beside : choose,
        };
        var answer = TaskDialog.ShowDialog(this, page);
        if (answer == beside)
        {
            _settings.SaveEachBesideSource = true;
            return true;
        }
        if (answer != choose) return false;

        folder = AskOutputFolder(entries);
        if (folder is null) return false;
        _settings.SaveEachBesideSource = false;
        return true;
    }

    /// <summary>ファイルごとの PDF を保存するフォルダを尋ねる。最初に開くのは、前回の保存先か、先頭のファイルのフォルダ。</summary>
    private string? AskOutputFolder(IReadOnlyList<FileEntry> entries)
    {
        string firstDir = Path.GetDirectoryName(entries[0].Path)!;
        using var dialog = new FolderBrowserDialog
        {
            Description = "ファイルごとの PDF を保存するフォルダ",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            InitialDirectory = Directory.Exists(_settings.LastOutputDirectory) ? _settings.LastOutputDirectory! : firstDir,
        };
        return dialog.ShowDialog(this) == DialogResult.OK ? Path.GetFullPath(dialog.SelectedPath) : null;
    }

    private static string ProblemList(IEnumerable<FileEntry> entries) =>
        string.Join("\r\n", entries.Select(e =>
        {
            string reason = e.State switch
            {
                EntryState.Ready => "指定したページがありません。",
                EntryState.NeedsPassword => e.Error ?? "パスワードが必要です。",
                _ => e.Error ?? "PDF にできませんでした。",
            };
            return $"・{e.FileName}\r\n    {reason}";
        }));
}
