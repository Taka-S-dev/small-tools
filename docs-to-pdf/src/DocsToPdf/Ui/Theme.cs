namespace DocsToPdf.Ui;

/// <summary>
/// 画面の色。Windows の「アプリ モード」（ライト / ダーク）に合わせる。
/// コントロールの標準の色は WinForms が切り替えるので、ここには自分で描く部分の色だけを置く。
/// PDF のページそのものは紙なので、どちらのモードでも白のまま見せる。
/// </summary>
internal static class Theme
{
    /// <summary>設定の値（"System" / "Light" / "Dark"）から、使う色のモードを決める。起動時、画面を作る前に呼ぶ。</summary>
    public static void Apply(string? setting)
    {
        var mode = setting?.Trim().ToLowerInvariant() switch
        {
            "light" => SystemColorMode.Classic,
            "dark" => SystemColorMode.Dark,
            _ => SystemColorMode.System,
        };
#pragma warning disable WFO5001 // ダークモードの API は .NET では「試験的」扱い。使い方は公式の案内どおり
        Application.SetColorMode(mode);
#pragma warning restore WFO5001
    }

#pragma warning disable WFO5001
    public static bool IsDark => Application.IsDarkModeEnabled;
#pragma warning restore WFO5001

    [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    /// <summary>
    /// タイトルバーも暗くする（Windows 10 20H1 以降）。ウィンドウの OnHandleCreated から呼ぶ。
    /// タイトルバーにアクセントカラーを使う設定のときは、Windows がその色を優先する。
    /// </summary>
    public static void ApplyTitleBar(Form form)
    {
        const int UseImmersiveDarkMode = 20;
        int dark = IsDark ? 1 : 0;
        try { DwmSetWindowAttribute(form.Handle, UseImmersiveDarkMode, ref dark, sizeof(int)); }
        catch { /* 古い Windows では何もしない */ }
    }

    private static Color Pick(Color light, Color dark) => IsDark ? dark : light;

    /// <summary>結合できない・エラーを表す文字の色。暗い背景でも読める明るさにする。</summary>
    public static Color Danger => Pick(Color.Firebrick, Color.FromArgb(255, 138, 128));

    /// <summary>リンクの文字の色。標準の青は暗い背景に沈むので、ダークでは明るい青にする。</summary>
    public static Color Link => Pick(Color.FromArgb(0, 102, 204), Color.FromArgb(110, 180, 255));

    /// <summary>プレビューで、ページの後ろに敷く色。</summary>
    public static Color PreviewBack => Pick(Color.FromArgb(236, 236, 236), Color.FromArgb(38, 38, 38));

    /// <summary>拡大表示で、ページの後ろに敷く色。</summary>
    public static Color ViewerBack => Pick(Color.FromArgb(96, 96, 96), Color.FromArgb(24, 24, 24));

    public static Color PageBorder => Pick(Color.FromArgb(190, 190, 190), Color.FromArgb(90, 90, 90));

    /// <summary>結合しないページに重ねる覆い。背景と同じ色を半透明にして、ページを沈ませる。</summary>
    public static Color Veil(Color back, int alpha) => Color.FromArgb(alpha, back);

    public static Color Cross => Pick(Color.FromArgb(150, 150, 150), Color.FromArgb(120, 120, 120));

    /// <summary>ページに重ねる小さなボタン。ページ（白）の上に乗るので、両モードとも明るい地にする。</summary>
    public static Color OverlayButton => Color.FromArgb(235, 255, 255, 255);
    public static Color OverlayButtonEdge => Color.FromArgb(160, 160, 160);
    public static Color OverlayButtonText => Color.FromArgb(32, 32, 32);
}
