using DocsToPdf.Core;

namespace DocsToPdf.Ui;

// 「PDF を作成」: 残っている変換を待ち、結合して保存する
internal sealed partial class MainForm
{
    /// <summary>
    /// 一覧の順に結合して保存する。
    /// </summary>
    /// <param name="outputPath">保存先。null なら尋ねる。指定があるのは、スクリプトから使ったとき（--out）。</param>
    /// <returns>PDF を作れたら true。</returns>
    private async Task<bool> CreateAsync(string? outputPath)
    {
        if (_busy || _list.Items.Count == 0) return false;
        EndEdit(commit: true);
        bool batch = outputPath is not null;

        // 保存先を選んでいる間も含めて、一覧を変えられないようにする（選んだ後で行が増減しないように）
        var cts = _cts = new CancellationTokenSource();
        SetBusy(true);
        ShowProgress(false);
        try
        {
            // 元のファイルが変わっていたら、変換し直す
            await _sourceCheck;
            await CheckSourcesAsync();

            var entries = Entries().ToList();
            string? output = outputPath ?? AskOutputPath(entries);
            if (output is null) return false;
            ShowProgress(true);

            // まだ終わっていない変換を待つ
            await WaitForConversionsAsync(entries, cts.Token);

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
                if (!ConfirmPartial(unusable, usable.Count))
                {
                    SetStatus("作成を取りやめました。");
                    return false;
                }
            }

            // 結合して保存する
            SetStatus("PDF を結合しています...");
            var parts = usable
                .Select(e => new MergePart(e.Converted!.PdfPath, e.Title, e.SelectedPages, e.Converted.Outline, e.Rotations))
                .ToList();
            var options = CurrentOptions(Path.GetFileNameWithoutExtension(output));
            _progress.Maximum = parts.Count;
            _progress.Value = 0;
            var progress = new Progress<int>(done =>
            {
                _progress.Value = Math.Min(_progress.Maximum, done);
                _taskbar.Set(Handle, done, parts.Count);
            });
            int pages = await Task.Run(() => PdfMerger.Merge(parts, options, output, progress, cts.Token), cts.Token);

            _settings.LastOutputDirectory = Path.GetDirectoryName(output);
            SetCreatedStatus(output, usable.Count, pages);
            WindowAttention.Flash(this);
            if (_optOpen.Checked && !_closeWhenStopped) ShellOpen(output);
            return true;
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

    private bool ConfirmPartial(IReadOnlyList<FileEntry> unusable, int usableCount)
    {
        var merge = new TaskDialogButton($"残りの {usableCount} 件で作成");
        var page = new TaskDialogPage
        {
            Caption = AppTitle,
            Heading = $"{unusable.Count} 件のファイルを結合できません",
            Text = "結合できるファイルだけで PDF を作りますか？",
            Icon = TaskDialogIcon.Warning,
            Expander = new TaskDialogExpander(ProblemList(unusable)) { Expanded = true, Position = TaskDialogExpanderPosition.AfterText },
            Buttons = { merge, TaskDialogButton.Cancel },
            DefaultButton = merge,
        };
        return TaskDialog.ShowDialog(this, page) == merge;
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
