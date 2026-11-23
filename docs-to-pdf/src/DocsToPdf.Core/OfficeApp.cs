using System.Diagnostics;
using System.Runtime.InteropServices;

namespace DocsToPdf.Core;

/// <summary>
/// このツールが使う Office 1 個。作ったのと同じ STA スレッドからだけ呼ぶこと。
/// 起動は最初の変換時に行い、以降のファイルでは使い回す。
/// 応答がなくなったときは、プロセスを強制終了して次のファイルで起動し直す。
///
/// 利用者の作業を巻き添えにしないための決まり:
/// ・終了や強制終了をするのは、このツールが起動したと確認できた Office だけ。
///   利用者がもともと開いていた Office につながった場合（PowerPoint は 1 つしか起動しない）は、閉じも止めもしない。
/// ・このツールが起動した見えない Office の中で、利用者があとから別のファイルを開くことがある
///   （Office は、すでに動いているものを使い回してファイルを開くため）。
///   それを見つけたら、その Office を画面に出して利用者に引き渡し、以後は触らない。
/// </summary>
internal abstract class OfficeApp : IDisposable
{
    /// <summary>パスワード付きファイルで入力を求められて止まらないよう、パスワードの指定がないときにわざと渡す誤った値。</summary>
    protected const string WrongPassword = "__no_password_given__";

    private readonly string _progId;
    private readonly string _processName;
    /// <summary>文書を出す窓のクラス名（Word の OpusApp など）。これが見えていれば、利用者が文書を開いている。</summary>
    private readonly string _documentWindowClass;
    private dynamic? _app;
    private volatile Process? _process;
    // 時間切れや確認の画面の見張り（別スレッド）からも読むので volatile
    private volatile bool _owned;
    private volatile bool _killed;
    /// <summary>確認の画面で止まって打ち切ったときの、その画面の文面。</summary>
    private volatile string? _stuckDialog;

    protected OfficeApp(string progId, string processName, string displayName, string documentWindowClass)
    {
        _progId = progId;
        _processName = processName;
        _documentWindowClass = documentWindowClass;
        DisplayName = displayName;
    }

    public string DisplayName { get; }

    protected abstract void Configure(dynamic app);

    /// <summary>
    /// 利用者の Office を借りているとき、変換の間だけ設定を変える（マクロを実行しない、確認を出さない）。
    /// 元に戻す処理を返す。変えられなかった設定があれば、変えた分を戻してから例外を投げる。
    /// </summary>
    protected abstract Action Borrow(dynamic app);

    /// <summary>Borrow で変えた設定と、その戻し方。戻すときは変えたのと逆の順に戻す。</summary>
    internal sealed class SettingChanges
    {
        private readonly List<Action> _undo = [];

        public void Change<T>(Func<T> get, Action<T> set, T value)
        {
            T original = get();
            set(value);
            _undo.Add(() => set(original));
        }

        public void Undo()
        {
            for (int i = _undo.Count - 1; i >= 0; i--)
            {
                try { _undo[i](); } catch { /* Office が閉じられた。残りは戻せるだけ戻す */ }
            }
            _undo.Clear();
        }

        /// <summary>changes で設定を変え、戻す処理を返す。途中で失敗したら、変えた分を戻してから投げる。</summary>
        public static Action Apply(Action<SettingChanges> changes)
        {
            var applied = new SettingChanges();
            try { changes(applied); }
            catch
            {
                applied.Undo();
                throw;
            }
            return applied.Undo;
        }
    }

    protected abstract void QuitCore(dynamic app);

    /// <summary>このツールが開いたもの以外に、利用者の目に見える形で開いている文書の数。</summary>
    protected abstract int ForeignDocumentCount(dynamic app);

    /// <summary>隠していた Office を画面に出し、止めていた確認メッセージなどを元に戻す。</summary>
    protected abstract void HandOver(dynamic app);

    /// <summary>
    /// つながった直後の Office が、今まさに自動操作のために起動されたものに見えるか（画面に出ておらず、何も開いていない）。
    /// 利用者が使っている Office は画面に出ているので、これで見分けられる。
    /// </summary>
    protected abstract bool LooksFresh(dynamic app);

    /// <summary>このツールが起動した Office か。false なら、利用者の Office を借りているだけ。</summary>
    protected bool Owned => _owned;

    /// <summary>
    /// つながった Office のプロセス ID。既定では起動前後のプロセス一覧の差で見つける。
    /// 同じ瞬間に利用者も Office を起動して差が複数になったときは、画面を持たないものを選ぶ。
    /// </summary>
    protected virtual int? FindProcessId(dynamic app, HashSet<int> before)
    {
        var added = ProcessIds().Except(before).ToList();
        if (added.Count == 1) return added[0];

        var hidden = new List<int>();
        foreach (int id in added)
        {
            try
            {
                using var process = Process.GetProcessById(id);
                if (process.MainWindowHandle == IntPtr.Zero) hidden.Add(id);
            }
            catch { /* もう終わっている */ }
        }
        return hidden.Count == 1 ? hidden[0] : null;
    }

    protected static int? ProcessIdOfWindow(int hwnd)
    {
        NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd), out uint pid);
        return pid != 0 ? (int)pid : null;
    }

    /// <summary>Office に対する処理を、時間切れと終了の求めを見張りながら実行する。</summary>
    protected T Run<T>(Func<dynamic, T> action, TimeSpan timeout, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        EnsureStarted();
        _killed = false;
        _stuckDialog = null;

        T result;
        using (new System.Threading.Timer(_ => Kill(onlyIfHidden: false), null, timeout, Timeout.InfiniteTimeSpan))
        using (ct.Register(() => Kill(onlyIfHidden: true)))
        // 自分で起動した Office が確認の画面を出したまま止まったら、時間切れを待たずに止める。
        // 利用者の Office を借りているときは、その画面が利用者のものかもしれないので見張らない
        using (_owned && _process is not null
                   ? new DialogWatch(_process, text => { _stuckDialog = text; Kill(onlyIfHidden: false); })
                   : null)
        {
            try
            {
                // 利用者の Office を借りているときは、マクロを無効にできなければ開かない。
                // 戻すのは、下の catch で接続を捨てるより前（ここの finally）
                Action? restore = null;
                if (!_owned)
                {
                    try { restore = Borrow(_app!); }
                    catch (Exception ex) when (!IsDisconnected(ex))
                    {
                        throw new InvalidOperationException($"{DisplayName} でマクロを無効にできないため、変換しませんでした。", ex);
                    }
                }
                try { result = action(_app!); }
                finally { restore?.Invoke(); }
            }
            catch (Exception) when (_killed)
            {
                Reset();
                ct.ThrowIfCancellationRequested();
                if (_stuckDialog is { } dialog)
                    throw new InvalidOperationException(
                        $"{DisplayName} が確認の画面を出したまま止まったため、中止しました" + (dialog.Length > 0 ? $"（{dialog}）。" : "。"));
                string limit = timeout.TotalMinutes >= 1 ? $"{timeout.TotalMinutes:0} 分" : $"{timeout.TotalSeconds:0.#} 秒";
                throw new TimeoutException($"{limit}たっても終わらないため打ち切りました（{DisplayName} が応答しません）。");
            }
            catch (Exception ex) when (_process is { HasExited: true } || IsDisconnected(ex))
            {
                // 利用者の Office を借りていて、その Office が閉じられた場合もここに来る。
                // 接続を捨てておけば、次のファイルでは起動し直せる
                Reset();
                throw new InvalidOperationException($"{DisplayName} が途中で終了しました。", ex);
            }
            catch
            {
                // 変換に失敗したファイルの間に利用者のファイルが開かれていた場合も、ここで引き渡す
                RelinquishIfInUse();
                throw;
            }
        }

        // 処理が終わった直後に時間切れや終了の求めが重なった場合
        if (_killed)
        {
            Reset();
            ct.ThrowIfCancellationRequested();
        }

        // 変換している間に利用者のファイルがこの Office で開かれていたら、ここで引き渡す
        RelinquishIfInUse();
        return result;
    }

    /// <summary>
    /// 利用者がこの Office を使い始めていたら、画面に出して引き渡し、こちらからは手を離す。
    /// </summary>
    /// <returns>引き渡したら true。</returns>
    private bool RelinquishIfInUse()
    {
        if (_app is null || !_owned) return false;
        bool inUse;
        try { inUse = ForeignDocumentCount(_app) > 0 || HasVisibleWindow(); }
        catch { inUse = HasVisibleWindow(); }
        if (!inUse) return false;

        try { HandOver(_app); } catch { /* 出せるところまで出す */ }
        _owned = false;
        ReleaseApp();
        _process?.Dispose();
        _process = null;
        return true;
    }

    /// <summary>相手の Office がもういない（終了した、接続が切れた）ことを表すエラーか。</summary>
    private static bool IsDisconnected(Exception ex) => ex is COMException com && unchecked((uint)com.HResult) is
        0x800706BA    // RPC サーバーを利用できない
        or 0x800706BE // リモート プロシージャ コールに失敗した
        or 0x80010108 // 起動されたオブジェクトはクライアントから切断された
        or 0x80010007; // サーバーが呼び出しの途中で終了した

    private bool HasVisibleWindow()
    {
        if (_process is null) return false;
        try
        {
            _process.Refresh();
            return !_process.HasExited && _process.MainWindowHandle != IntPtr.Zero;
        }
        catch
        {
            return false;
        }
    }

    private void EnsureStarted()
    {
        if (_app is not null && (_process is null || !_process.HasExited))
            return;
        Reset();

        var type = Type.GetTypeFromProgID(_progId)
            ?? throw new InvalidOperationException($"{DisplayName} がインストールされていません。");

        var before = ProcessIds();
        _app = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException($"{DisplayName} を起動できませんでした。");
        try
        {
            int? pid = FindProcessId(_app, before);
            _process = pid is int id ? Process.GetProcessById(id) : null;
            // 起動前からあったプロセスや、すでに画面に出ている Office は、利用者のものにつながっただけ。
            // プロセスを特定できなかった場合も、起動したての見た目なら自分のものとして終了まで面倒を見る
            // （その場合は COM 経由の終了だけで、強制終了はできない）
            bool existedBefore = pid is int known && before.Contains(known);
            _owned = !existedBefore && LooksFresh(_app);
            if (_owned && _process is not null) OfficeLedger.Record(_process);
            // 借りているだけの Office の設定は、ここでは変えない（変換のたびに変えて戻す）
            if (_owned) Configure(_app);
        }
        catch
        {
            Reset();
            throw;
        }
    }

    /// <summary>タイマー・確認の画面の見張り・終了処理（別スレッド）から呼ばれる。</summary>
    /// <param name="onlyIfHidden">
    /// 画面に出ていない Office だけを止める。アプリを閉じるときに使う。
    /// 画面に出ているなら利用者のファイルが開かれているので、止めずに残す。
    /// </param>
    private void Kill(bool onlyIfHidden)
    {
        var process = _process;
        if (!MayKill(_owned, onlyIfHidden, HasVisibleWindow(), ShowsUserDocument(process))) return;
        _killed = true;
        try { process?.Kill(entireProcessTree: true); } catch { /* もう終わっている */ }
    }

    /// <summary>
    /// 自分で起動した Office を止めてよいか。
    /// 変換している間に、利用者がこの Office で自分の文書を開いていることがある（Office は動いているものを使い回して開くため）。
    /// その文書の窓が見えているときは、時間切れや確認の画面で止まったときでも止めない（編集中の文書ごと消えてしまう）。
    /// 確認の画面も、利用者の文書が出したもの（「保存しますか」など）かもしれない。
    /// 止めなければ、Office が返ってきたところで、利用者に引き渡す（RelinquishIfInUse）。
    /// </summary>
    internal static bool MayKill(bool owned, bool onlyIfHidden, bool anyWindowVisible, bool userDocumentVisible) =>
        owned && !userDocumentVisible && !(onlyIfHidden && anyWindowVisible);

    /// <summary>利用者の文書の窓が見えているか。COM を使わずに窓を数えるので、変換で止まっていても、ほかのスレッドから調べられる。</summary>
    private bool ShowsUserDocument(Process? process)
    {
        try
        {
            return process is { HasExited: false }
                && DialogWatch.FindVisibleWindow(process.Id, name => name == _documentWindowClass) != IntPtr.Zero;
        }
        catch
        {
            return false; // もう終わっている
        }
    }

    /// <summary>今の Office との接続を捨てる。自分で起動したものが生きていれば強制終了する。</summary>
    private void Reset()
    {
        ReleaseApp();
        if (_process is null) return;
        if (_owned)
        {
            try
            {
                if (!_process.HasExited)
                {
                    _process.Kill(entireProcessTree: true);
                    _process.WaitForExit(5_000);
                }
            }
            catch { /* もう終わっている */ }
        }
        _process.Dispose();
        _process = null;
        _owned = false;
    }

    /// <summary>
    /// 終了を頼むだけで、終わるのは待たない。複数の Office をまとめて終わらせるとき、
    /// 先に全部へ頼んでおくと待ち時間が重ならない。このあと Dispose を呼ぶこと。
    /// </summary>
    public void BeginQuit()
    {
        if (_app is null) return;
        if (RelinquishIfInUse()) return;
        if (_owned)
        {
            try { QuitCore(_app); } catch { /* Dispose で強制終了する */ }
        }
        ReleaseApp();
    }

    public void Dispose()
    {
        BeginQuit();
        if (_process is null) return;
        if (_owned)
        {
            try
            {
                // 開いた文書は保存しない写しだけなので、少し待って終わらなければ強制終了してよい
                if (!_process.WaitForExit(4_000))
                    _process.Kill(entireProcessTree: true);
            }
            catch { /* もう終わっている */ }
        }
        _process.Dispose();
        _process = null;
        _owned = false;
    }

    private void ReleaseApp()
    {
        if (_app is null) return;
        try { Marshal.FinalReleaseComObject(_app); } catch { /* 接続が切れている */ }
        _app = null;
        // dynamic 経由で作られた中間の COM 参照も解放し、Office が終われるようにする
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    protected static void Release(object? comObject)
    {
        if (comObject is null || !Marshal.IsComObject(comObject)) return;
        try { Marshal.FinalReleaseComObject(comObject); } catch { /* 接続が切れている */ }
    }

    /// <summary>
    /// 「保存時に個人情報を削除する」が設定された文書は、保存や PDF への書き出しのたびに
    /// 「ドキュメント検査ですべては削除できません」と警告し、確認を出さない設定にしていても答えを待って止まる。
    /// 開いたのは作業フォルダの写しなので、その写しの設定だけを切る（元のファイルは変わらない）。
    /// </summary>
    protected static void KeepPersonalInformation(object document)
    {
        try
        {
            dynamic doc = document;
            if ((bool)doc.RemovePersonalInformation) doc.RemovePersonalInformation = false;
        }
        catch { /* 設定のない古い Office や、変えられない文書。そのときは警告も出ない */ }
    }

    private HashSet<int> ProcessIds()
    {
        var ids = new HashSet<int>();
        foreach (var p in Process.GetProcessesByName(_processName))
        {
            ids.Add(p.Id);
            p.Dispose();
        }
        return ids;
    }
}

internal sealed class WordApp : OfficeApp
{
    public WordApp() : base("Word.Application", "WINWORD", "Word", "OpusApp") { }

    protected override void Configure(dynamic app)
    {
        app.Visible = false;
        app.ScreenUpdating = false;
        app.DisplayAlerts = 0;      // wdAlertsNone
        app.AutomationSecurity = 3; // msoAutomationSecurityForceDisable: マクロを実行しない
    }

    protected override Action Borrow(dynamic app) => SettingChanges.Apply(c =>
    {
        c.Change<int>(() => (int)app.DisplayAlerts, v => app.DisplayAlerts = v, 0);           // wdAlertsNone
        c.Change<int>(() => (int)app.AutomationSecurity, v => app.AutomationSecurity = v, 3); // msoAutomationSecurityForceDisable
    });

    /// <summary>文書を PDF にする。見出しはしおりになる。</summary>
    /// <param name="small">「最小サイズ」で書き出す（画像の解像度を下げる）。</param>
    public void Export(string source, string pdf, string? password, bool small, TimeSpan timeout, CancellationToken ct) =>
        Run<object?>(app =>
        {
            dynamic documents = app.Documents;
            dynamic? doc = null;
            try
            {
                // ほかのファイルへのリンクを含む文書は、確認を出さない設定にしていても「リンクを更新しますか」と聞いて止まる。
                // 開く間だけ更新を止める。これは利用者の Word 全体の設定なので、開いたらすぐ元に戻す
                bool? updateLinks = null;
                try
                {
                    updateLinks = (bool)app.Options.UpdateLinksAtOpen;
                    app.Options.UpdateLinksAtOpen = false;
                }
                catch { /* 設定を読めなければ、そのまま開く */ }
                try
                {
                    doc = documents.Open(
                        FileName: source,
                        ConfirmConversions: false,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        PasswordDocument: password ?? WrongPassword,
                        PasswordTemplate: password ?? WrongPassword,
                        Revert: false,
                        Visible: false,
                        OpenAndRepair: false,
                        NoEncodingDialog: true);
                }
                finally
                {
                    if (updateLinks is bool original)
                    {
                        try { app.Options.UpdateLinksAtOpen = original; } catch { /* Word が応答しない（このあと打ち切る） */ }
                    }
                }

                KeepPersonalInformation(doc);
                HideMarkup(doc);

                doc.ExportAsFixedFormat(
                    OutputFileName: pdf,
                    ExportFormat: 17,        // wdExportFormatPDF
                    OpenAfterExport: false,
                    OptimizeFor: small ? 1 : 0, // wdExportOptimizeForOnScreen / wdExportOptimizeForPrint
                    Range: 0,                // wdExportAllDocument
                    Item: 0,                 // wdExportDocumentContent
                    IncludeDocProps: true,
                    KeepIRM: true,
                    CreateBookmarks: 1,      // wdExportCreateHeadingBookmarks
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false);
            }
            finally
            {
                if (doc is not null)
                {
                    try { doc.Close(SaveChanges: 0); } catch { /* 閉じられなくても続ける */ }
                    Release(doc);
                }
                Release(documents);
            }
            return null;
        }, timeout, ct);

    /// <summary>変更履歴とコメントを出さず、変更を反映した最終版の見た目で出力する。</summary>
    private static void HideMarkup(dynamic doc)
    {
        try
        {
            dynamic filter = doc.ActiveWindow.View.RevisionsFilter;
            filter.Markup = 0; // wdRevisionsMarkupNone
            filter.View = 0;   // wdRevisionsViewFinal
        }
        catch { /* 古い Word などでは既定の表示のまま出す */ }
    }

    protected override void QuitCore(dynamic app) => app.Quit(SaveChanges: 0);

    protected override int ForeignDocumentCount(dynamic app) => (int)app.Documents.Count;

    protected override bool LooksFresh(dynamic app) => !(bool)app.Visible && (int)app.Documents.Count == 0;

    protected override void HandOver(dynamic app)
    {
        try { app.DisplayAlerts = -1; } catch { }      // wdAlertsAll
        try { app.AutomationSecurity = 1; } catch { }  // msoAutomationSecurityLow（既定）
        try { app.ScreenUpdating = true; } catch { }
        app.Visible = true;
    }
}

internal sealed class ExcelApp : OfficeApp
{
    private const int SheetVisible = -1; // xlSheetVisible

    public ExcelApp() : base("Excel.Application", "EXCEL", "Excel", "XLMAIN") { }

    protected override void Configure(dynamic app)
    {
        app.Visible = false;
        app.ScreenUpdating = false;
        app.DisplayAlerts = false;
        app.EnableEvents = false;
        app.AskToUpdateLinks = false;
        app.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
    }

    protected override Action Borrow(dynamic app) => SettingChanges.Apply(c =>
    {
        c.Change<bool>(() => (bool)app.DisplayAlerts, v => app.DisplayAlerts = v, false);
        c.Change<bool>(() => (bool)app.EnableEvents, v => app.EnableEvents = v, false);
        c.Change<int>(() => (int)app.AutomationSecurity, v => app.AutomationSecurity = v, 3); // msoAutomationSecurityForceDisable
    });

    protected override int? FindProcessId(dynamic app, HashSet<int> before)
    {
        int hwnd = app.Hwnd;
        return ProcessIdOfWindow(hwnd) ?? base.FindProcessId((object)app, before);
    }

    /// <summary>
    /// ブックを PDF にする。非表示のシートは出ず、印刷範囲があるシートはその範囲だけ出る。
    /// </summary>
    /// <param name="sheets">出力するシート名。null なら表示されているシートすべて。</param>
    /// <param name="fitWidth">出力するシートのすべての列を、1 ページの幅に収める。</param>
    /// <param name="small">「最小サイズ」で書き出す（画像の解像度を下げる）。</param>
    /// <returns>表示されているシートの一覧と、それぞれのページ数（数えられなければ 0）。</returns>
    public IReadOnlyList<SheetInfo> Export(string source, string pdf, string? password,
        IReadOnlyList<string>? sheets, bool fitWidth, bool small, TimeSpan timeout, CancellationToken ct) =>
        Run(app =>
        {
            dynamic workbooks = app.Workbooks;
            dynamic? book = null;
            try
            {
                book = workbooks.Open(
                    Filename: source,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Password: password ?? WrongPassword,
                    IgnoreReadOnlyRecommended: true,
                    Notify: false,
                    AddToMru: false);
                KeepPersonalInformation(book);

                var visible = new List<(string Name, object Sheet)>();
                dynamic all = book.Sheets;
                int count = all.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic sheet = all[i];
                    if ((int)sheet.Visible == SheetVisible)
                        visible.Add(((string)sheet.Name, (object)sheet));
                }

                var wanted = sheets is null ? null : new HashSet<string>(sheets, StringComparer.Ordinal);
                var included = visible.Where(v => wanted is null || wanted.Contains(v.Name)).ToList();
                if (included.Count == 0)
                    throw new InvalidOperationException("出力するシートがありません（選んだシートがブックに見つかりません）。");

                if (fitWidth)
                    FitToOnePageWide((object)app, included.Select(v => v.Sheet).ToList());

                var result = new List<SheetInfo>(visible.Count);
                foreach (var (name, sheet) in visible)
                {
                    bool isIncluded = wanted is null || wanted.Contains(name);
                    result.Add(new SheetInfo(name, isIncluded ? CountPages(sheet) : 0, isIncluded));
                }

                if (included.Count == visible.Count)
                {
                    book.ExportAsFixedFormat(
                        Type: 0,                 // xlTypePDF
                        Filename: pdf,
                        Quality: small ? 1 : 0,  // xlQualityMinimum / xlQualityStandard
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);
                }
                else
                {
                    // 選んだシートだけをまとめて選択し、その選択を出力する
                    bool replace = true;
                    foreach (var (_, sheet) in included)
                    {
                        ((dynamic)sheet).Select(replace);
                        replace = false;
                    }
                    dynamic active = app.ActiveSheet;
                    active.ExportAsFixedFormat(
                        Type: 0,
                        Filename: pdf,
                        Quality: small ? 1 : 0,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);
                }
                return (IReadOnlyList<SheetInfo>)result;
            }
            finally
            {
                if (book is not null)
                {
                    try { book.Close(SaveChanges: false); } catch { /* 閉じられなくても続ける */ }
                    Release(book);
                }
                Release(workbooks);
            }
        }, timeout, ct);

    private static void FitToOnePageWide(object application, IReadOnlyList<object> sheets)
    {
        dynamic app = application;
        // ページ設定はプリンターとのやりとりで遅いので、まとめて反映させる
        try { app.PrintCommunication = false; } catch { /* 古い Excel にはない */ }
        try
        {
            foreach (dynamic sheet in sheets)
            {
                try
                {
                    dynamic setup = sheet.PageSetup;
                    setup.Zoom = false;
                    setup.FitToPagesWide = 1;
                    setup.FitToPagesTall = false;
                }
                catch { /* グラフシートなど、幅の指定がないシートはそのまま出す */ }
            }
        }
        finally
        {
            try { app.PrintCommunication = true; } catch { /* 古い Excel にはない */ }
        }
    }

    private static int CountPages(object sheet)
    {
        try { return (int)((dynamic)sheet).PageSetup.Pages.Count; }
        catch { return 0; }
    }

    protected override void QuitCore(dynamic app) => app.Quit();

    protected override bool LooksFresh(dynamic app) => !(bool)app.Visible;

    /// <summary>ウィンドウが表示されているブックを数える（個人用マクロブックのような隠れたブックは数えない）。</summary>
    protected override int ForeignDocumentCount(dynamic app)
    {
        int count = 0;
        dynamic books = app.Workbooks;
        int total = books.Count;
        for (int i = 1; i <= total; i++)
        {
            dynamic windows = books[i].Windows;
            if ((int)windows.Count > 0 && (bool)windows[1].Visible) count++;
        }
        return count;
    }

    protected override void HandOver(dynamic app)
    {
        try { app.DisplayAlerts = true; } catch { }
        try { app.EnableEvents = true; } catch { }
        try { app.AskToUpdateLinks = true; } catch { }
        try { app.AutomationSecurity = 1; } catch { }  // msoAutomationSecurityLow（既定）
        try { app.ScreenUpdating = true; } catch { }
        app.Visible = true;
        try { app.UserControl = true; } catch { }
    }
}

internal sealed class PowerPointApp : OfficeApp
{
    public PowerPointApp() : base("PowerPoint.Application", "POWERPNT", "PowerPoint", "PPTFrameClass") { }

    protected override void Configure(dynamic app)
    {
        // PowerPoint はウィンドウを隠す指定（Visible = false）を受け付けない。
        // ウィンドウなしで開けば画面には出ない。
        app.DisplayAlerts = 1;      // ppAlertsNone
        app.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
    }

    protected override Action Borrow(dynamic app) => SettingChanges.Apply(c =>
    {
        c.Change<int>(() => (int)app.DisplayAlerts, v => app.DisplayAlerts = v, 1);           // ppAlertsNone
        c.Change<int>(() => (int)app.AutomationSecurity, v => app.AutomationSecurity = v, 3); // msoAutomationSecurityForceDisable
    });

    // PowerPoint は 1 つしか起動しないので、利用者がすでに開いていれば、その PowerPoint につながる。
    // その場合は画面に出ているので「自分で起動したものではない」と判断し、終了も強制終了もしない。
    protected override bool LooksFresh(dynamic app) => (int)app.Visible == 0 && (int)app.Presentations.Count == 0;

    /// <summary>
    /// スライドを PDF にする。非表示のスライドは出ず、スライドのタイトルがしおりになる。
    /// 「最小サイズ」の書き出し（ExportAsFixedFormat）は、型ライブラリなしの呼び出しでは PowerPoint が受け付けない
    /// （引数の型が合わないと断られる）ので、小さく作るときは書き出したあとに PdfImages で縮める。
    /// </summary>
    public void Export(string source, string pdf, string? password, TimeSpan timeout, CancellationToken ct) =>
        Run<object?>(app =>
        {
            dynamic presentations = app.Presentations;
            dynamic? presentation = null;
            try
            {
                // "パス::パスワード::" と書くと、入力を求められずにパスワードを渡せる
                presentation = presentations.Open(
                    FileName: $"{source}::{password ?? WrongPassword}::",
                    ReadOnly: -1,    // msoTrue
                    Untitled: 0,     // msoFalse
                    WithWindow: 0);  // msoFalse
                KeepPersonalInformation(presentation); // PDF への変換は「保存」なので、この警告に当たる
                presentation.SaveAs(FileName: pdf, FileFormat: 32); // ppSaveAsPDF
            }
            finally
            {
                if (presentation is not null)
                {
                    try { presentation.Close(); } catch { /* 閉じられなくても続ける */ }
                    Release(presentation);
                }
                Release(presentations);
            }
            return null;
        }, timeout, ct);

    protected override void QuitCore(dynamic app) => app.Quit();

    /// <summary>ウィンドウを持つプレゼンテーションを数える（このツールはウィンドウなしで開く）。</summary>
    protected override int ForeignDocumentCount(dynamic app)
    {
        int count = 0;
        dynamic presentations = app.Presentations;
        int total = presentations.Count;
        for (int i = 1; i <= total; i++)
        {
            if ((int)presentations[i].Windows.Count > 0) count++;
        }
        return count;
    }

    protected override void HandOver(dynamic app)
    {
        try { app.DisplayAlerts = 2; } catch { }       // ppAlertsAll
        try { app.AutomationSecurity = 1; } catch { }  // msoAutomationSecurityLow（既定）
        app.Visible = -1;                              // msoTrue
    }
}
