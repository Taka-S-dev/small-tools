using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Core;

/// <summary>
/// ファイルを PDF にして作業フォルダに置く。アプリが動いている間ずっと使い回す。
/// Word / Excel / PowerPoint は COM が STA スレッドを必要とするので、専用のスレッドで 1 件ずつ順に処理する。
/// PDF と画像は Office を使わないので、別のスレッドで並行して処理する。
/// 依頼が途切れたらすぐ Office を終了させ、次の依頼でまた起動する。
/// </summary>
internal sealed class ConversionService : IDisposable
{
    /// <summary>
    /// 最後の変換からこの時間たったら Office を終了させる。
    /// 見えない Office が動いている間に利用者が別のファイルを開くと、その Office の中で開かれてしまうことがあるので、
    /// 続けて落とされたファイルをまとめて処理できる程度の短さにとどめる。
    /// </summary>
    private static readonly TimeSpan IdleBeforeQuit = TimeSpan.FromSeconds(2);

    private readonly BlockingCollection<Job> _queue = [];
    private readonly CancellationTokenSource _shutdown = new();
    private readonly SemaphoreSlim _lightSlots = new(2);
    private readonly Thread _thread;
    private readonly string _workDir;
    private readonly TimeSpan _timeout;
    private int _nextId;

    public ConversionService(string workDir, TimeSpan? timeoutPerFile = null)
    {
        _workDir = workDir;
        OfficeLedger.Use(workDir);
        _timeout = timeoutPerFile ?? TimeSpan.FromMinutes(5);
        _thread = new Thread(Work) { IsBackground = true, Name = "Office to PDF" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    /// <summary>
    /// 変換を頼む。失敗は例外ではなく結果（Error）で返る。
    /// cancel は「まだ始まっていなければやめる」ための合図で、始まった変換は最後まで行う。
    /// </summary>
    /// <param name="onStarted">変換が始まったときに、変換側のスレッドから呼ばれる。</param>
    public Task<ConvertOutcome> ConvertAsync(ConvertRequest request, Action? onStarted, CancellationToken cancel)
    {
        int id = Interlocked.Increment(ref _nextId);
        if (request.Kind is FileKind.Pdf or FileKind.Image)
            return ConvertLightAsync(request, id, onStarted, cancel);

        var job = new Job(request, id, onStarted, cancel);
        try
        {
            _queue.Add(job);
        }
        catch (InvalidOperationException)
        {
            job.Completion.TrySetCanceled();
        }
        return job.Completion.Task;
    }

    private async Task<ConvertOutcome> ConvertLightAsync(ConvertRequest request, int id, Action? onStarted, CancellationToken cancel)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancel, _shutdown.Token);
        await _lightSlots.WaitAsync(linked.Token).ConfigureAwait(false);
        try
        {
            return await Task.Run(() =>
            {
                linked.Token.ThrowIfCancellationRequested();
                onStarted?.Invoke();
                return Guarded(request, stamp => request.Kind == FileKind.Pdf
                    ? PreparePdf(request, id, stamp)
                    : PrepareImage(request, id, stamp));
            }, linked.Token).ConfigureAwait(false);
        }
        finally
        {
            _lightSlots.Release();
        }
    }

    private sealed record Job(ConvertRequest Request, int Id, Action? OnStarted, CancellationToken Cancel)
    {
        public TaskCompletionSource<ConvertOutcome> Completion { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    }

    private void Work()
    {
        using var filter = OleMessageFilter.Register();
        WordApp? word = null;
        ExcelApp? excel = null;
        PowerPointApp? powerPoint = null;

        void QuitAll()
        {
            word?.BeginQuit();
            excel?.BeginQuit();
            powerPoint?.BeginQuit();
            word?.Dispose();
            excel?.Dispose();
            powerPoint?.Dispose();
            word = null;
            excel = null;
            powerPoint = null;
        }

        try
        {
            while (!_shutdown.IsCancellationRequested)
            {
                Job? job;
                try
                {
                    if (!_queue.TryTake(out job, (int)IdleBeforeQuit.TotalMilliseconds, _shutdown.Token))
                    {
                        QuitAll();
                        continue;
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }

                if (job.Cancel.IsCancellationRequested)
                {
                    job.Completion.TrySetCanceled();
                    continue;
                }

                try
                {
                    job.OnStarted?.Invoke();
                    var outcome = Guarded(job.Request, stamp =>
                    {
                        var request = job.Request;
                        string copy = CopySource(request.Path, job.Id);
                        string pdf = PdfPathFor(job.Id);
                        try
                        {
                            IReadOnlyList<SheetInfo>? sheets = null;
                            switch (request.Kind)
                            {
                                case FileKind.Word:
                                    (word ??= new WordApp()).Export(copy, pdf, request.Password, _timeout, _shutdown.Token);
                                    break;
                                case FileKind.Excel:
                                    sheets = (excel ??= new ExcelApp()).Export(copy, pdf, request.Password,
                                        request.Sheets, request.FitWidth, _timeout, _shutdown.Token);
                                    break;
                                default:
                                    (powerPoint ??= new PowerPointApp()).Export(copy, pdf, request.Password, _timeout, _shutdown.Token);
                                    break;
                            }

                            if (!File.Exists(pdf))
                                throw new InvalidOperationException("PDF が作られませんでした。");
                            return Finish(pdf, stamp, sheets);
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException
                                                   && LooksLikePasswordProblem(ex, copy, passwordGiven: request.Password is not null))
                        {
                            throw new PasswordRequiredException();
                        }
                        finally
                        {
                            TryDelete(copy);
                        }
                    });
                    job.Completion.TrySetResult(outcome);
                }
                catch (OperationCanceledException)
                {
                    job.Completion.TrySetCanceled();
                }
                catch (Exception ex)
                {
                    job.Completion.TrySetResult(ConvertOutcome.Failure(Describe(ex), null));
                }
            }
        }
        finally
        {
            QuitAll();
            _queue.CompleteAdding();
            while (_queue.TryTake(out var rest))
                rest.Completion.TrySetCanceled();
        }
    }

    private sealed class PasswordRequiredException : Exception;

    /// <summary>失敗を結果に直す。終了の求め（OperationCanceledException）だけはそのまま投げる。</summary>
    private static ConvertOutcome Guarded(ConvertRequest request, Func<SourceStamp, ConvertOutcome> convert)
    {
        var stamp = SourceStamp.Of(request.Path);
        if (stamp is null)
            return ConvertOutcome.Failure("ファイルが見つかりません。", null);
        try
        {
            return convert(stamp.Value);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (PasswordRequiredException)
        {
            string message = request.Password is null ? "パスワードが必要です。" : "パスワードが違います。";
            return ConvertOutcome.Failure(message, stamp, needsPassword: true);
        }
        catch (Exception ex)
        {
            return ConvertOutcome.Failure(Describe(ex), stamp);
        }
    }

    private ConvertOutcome PreparePdf(ConvertRequest request, int id, SourceStamp stamp)
    {
        string pdf = PdfPathFor(id);
        PdfDocument source;
        try
        {
            source = request.Password is null
                ? PdfReader.Open(request.Path, PdfDocumentOpenMode.Import)
                : PdfReader.Open(request.Path, request.Password, PdfDocumentOpenMode.Import);
        }
        catch (PdfReaderException ex) when (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase)
                                             || ex.Message.Contains("protected", StringComparison.OrdinalIgnoreCase))
        {
            throw new PasswordRequiredException();
        }

        using (source)
        {
            if (source.PageCount == 0)
                throw new InvalidOperationException("ページがありません。");
            var (sizes, outline) = PdfInfo.Read(source);

            if (source.SecuritySettings.IsEncrypted)
            {
                // 保護された PDF は、保護を外した写しを作業フォルダに作る（プレビューと結合で使う）
                using var plain = new PdfDocument();
                foreach (var page in source.Pages) plain.AddPage(page);
                plain.Save(pdf);
            }
            else
            {
                // 元のファイルをつかんだままにしないよう、写しを使う
                File.Copy(request.Path, pdf, overwrite: true);
                File.SetAttributes(pdf, FileAttributes.Normal);
            }
            return ConvertOutcome.Success(new ConvertedPdf(pdf, sizes, outline, null), stamp);
        }
    }

    private ConvertOutcome PrepareImage(ConvertRequest request, int id, SourceStamp stamp)
    {
        string pdf = PdfPathFor(id);
        ImagePdf.Create(request.Path, pdf);
        return Finish(pdf, stamp, null);
    }

    private static ConvertOutcome Finish(string pdf, SourceStamp stamp, IReadOnlyList<SheetInfo>? sheets)
    {
        var (sizes, outline) = PdfInfo.Read(pdf);
        if (sizes.Count == 0)
            throw new InvalidOperationException("出力するページがありません。");

        if (sheets is not null)
        {
            // Excel はしおりを作らないので、シートごとのページ数からシート名のしおりを組み立てる。
            // 数えたページ数の合計が実際の PDF と合わないときは、ずれたしおりを付けないよう諦める。
            var printed = sheets.Where(s => s.Included && s.Pages > 0).ToList();
            if (printed.Sum(s => s.Pages) == sizes.Count)
            {
                var nodes = new List<OutlineNode>();
                int page = 0;
                foreach (var sheet in printed)
                {
                    nodes.Add(new OutlineNode(sheet.Name, page, []));
                    page += sheet.Pages;
                }
                outline = nodes.Count > 1 ? nodes : [];
            }
            else
            {
                sheets = sheets.Select(s => s with { Pages = 0 }).ToList();
            }
        }
        return ConvertOutcome.Success(new ConvertedPdf(pdf, sizes, outline, sheets), stamp);
    }

    /// <summary>
    /// 元のファイルを作業フォルダへコピーする。Office にはこの写しを開かせる。
    /// 利用者が開いているファイルのロック、長いパスや記号を含むパス、OneDrive 上のファイルで失敗しにくくなり、
    /// 元のファイルを Office に触らせずに済む。
    ///
    /// マクロ付きの形式（.docm .xlsm .pptm など）は、マクロを取り除いた写しにする（MacroStripper）。
    /// 暗号化されたファイルは zip ではないので取り除けず、そのまま写す。その場合も Office はマクロを強制的に無効にして開く。
    /// zip なのに中身が読めないときは、取り除けないまま開くことになるので変換しない。
    ///
    /// インターネットから来たファイルの印（Zone.Identifier）は写しにも残す。消すと安全側の判断を弱めることになり、
    /// 印が付いたままでも、確かめた環境では自動操作で変換できた（Word・Excel・PowerPoint）。
    /// </summary>
    private string CopySource(string path, int id)
    {
        string copy = Path.Combine(_workDir, $"{id:D5}{Path.GetExtension(path)}");
        bool stripped = false;
        if (MacroStripper.CanContainMacros(path))
        {
            try
            {
                MacroStripper.Strip(path, copy);
                stripped = true;
            }
            catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException or IOException)
            {
                if (!IsOleCompound(path))
                    throw new InvalidOperationException("ファイルの中身が読めず、マクロを取り除けないため、変換しませんでした。", ex);
                // 暗号化されたファイル。下でそのまま写す
            }
        }
        if (!stripped) File.Copy(path, copy, overwrite: true);
        File.SetAttributes(copy, FileAttributes.Normal);
        return copy;
    }

    private string PdfPathFor(int id) => Path.Combine(_workDir, $"{id:D5}.pdf");

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* 作業フォルダごと後で消す */ }
    }

    /// <summary>
    /// 失敗の原因がパスワードかを見分ける。暗号化された docx / xlsx / pptx は ZIP ではなく
    /// OLE 複合ファイル（先頭が D0 CF 11 E0）になるので、それで判定する。古い形式はエラーの内容で判定する。
    /// </summary>
    private static bool LooksLikePasswordProblem(Exception ex, string file, bool passwordGiven)
    {
        if (ex is TimeoutException) return false;
        if (ex is COMException { HResult: unchecked((int)0x800A1520) }) return true; // Word: パスワードが正しくない
        if (ex.Message.Contains("パスワード") || ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
            return true;
        // パスワードを渡したうえでの失敗は、内容からパスワードが原因と分かるものだけを扱う
        if (passwordGiven) return false;

        string ext = Path.GetExtension(file).ToLowerInvariant();
        return ext is ".docx" or ".docm" or ".xlsx" or ".xlsm" or ".xlsb" or ".pptx" or ".pptm" or ".ppsx"
            && IsOleCompound(file);
    }

    /// <summary>OLE 複合ファイル（先頭が D0 CF 11 E0）か。暗号化された docx / xlsx / pptx はこの形になる。</summary>
    internal static bool IsOleCompound(string file)
    {
        try
        {
            Span<byte> head = stackalloc byte[4];
            using var stream = File.OpenRead(file);
            return stream.Read(head) == 4 && head.SequenceEqual<byte>([0xD0, 0xCF, 0x11, 0xE0]);
        }
        catch
        {
            return false;
        }
    }

    private static string Describe(Exception ex)
    {
        string message = ex switch
        {
            COMException com => $"{FirstLine(com.Message)} (0x{com.HResult:X8})",
            _ => FirstLine(ex.Message),
        };
        return string.IsNullOrWhiteSpace(message) ? ex.GetType().Name : message;
    }

    private static string FirstLine(string text) =>
        (text.ReplaceLineEndings("\n").Split('\n', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "").Trim();

    /// <summary>変換をやめて Office を終了させる。変換中なら Office を強制終了して待たない。</summary>
    public void Dispose()
    {
        if (_shutdown.IsCancellationRequested) return;
        _shutdown.Cancel();
        _thread.Join(TimeSpan.FromSeconds(20));
    }
}
