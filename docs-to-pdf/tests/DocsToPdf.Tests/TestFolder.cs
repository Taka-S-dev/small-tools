using PdfSharp.Pdf;

namespace DocsToPdf.Tests;

/// <summary>テストで使う PDF を作る。作ったファイルは Dispose でフォルダごと消える。</summary>
internal sealed class TestFolder : IDisposable
{
    public string Path { get; } = Directory.CreateTempSubdirectory("DocsToPdf.Tests.").FullName;

    public string File(string name) => System.IO.Path.Combine(Path, name);

    /// <param name="outline">しおり: (名前, 0 始まりのページ) の一覧。</param>
    public string Pdf(string name, int pages, params (string Title, int Page)[] outline)
    {
        string path = File(name);
        using var doc = new PdfDocument();
        for (int i = 0; i < pages; i++) doc.AddPage();
        foreach (var (title, page) in outline) doc.Outlines.Add(title, doc.Pages[page]);
        doc.Save(path);
        return path;
    }

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
    }
}
