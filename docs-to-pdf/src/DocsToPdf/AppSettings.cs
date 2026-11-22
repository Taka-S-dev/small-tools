using System.Text.Json;

namespace DocsToPdf;

/// <summary>前回の保存先、オプション、ウィンドウの位置。%LOCALAPPDATA%\DocsToPdf\settings.json に置く。</summary>
internal sealed class AppSettings
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "DocsToPdf", "settings.json");

    public string? LastOutputDirectory { get; set; }
    public bool PageNumbers { get; set; }
    public bool TableOfContents { get; set; }
    public bool DuplexBlanks { get; set; }
    public bool OpenAfterCreate { get; set; } = true;
    /// <summary>「個別に PDF を保存」で、前回それぞれ元のファイルと同じフォルダに保存したか（次回の既定にする）。</summary>
    public bool SaveEachBesideSource { get; set; } = true;
    /// <summary>画面の色: "System"（Windows の設定に合わせる）/ "Light" / "Dark"。</summary>
    public string ColorMode { get; set; } = "System";

    public int WindowX { get; set; }
    public int WindowY { get; set; }
    public int WindowWidth { get; set; }
    public int WindowHeight { get; set; }
    public bool Maximized { get; set; }
    /// <summary>一覧とプレビューの境目の位置（一覧側の幅の割合）。</summary>
    public double SplitRatio { get; set; } = 0.62;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(FilePath)) ?? new AppSettings();
        }
        catch { /* 壊れていたら既定値 */ }
        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { /* 保存できなくても動作には影響しない */ }
    }
}
