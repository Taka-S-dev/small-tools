using System.IO.Compression;
using System.Text;
using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class MacroStripperTests : IDisposable
{
    private readonly TestFolder _folder = new();

    public void Dispose() => _folder.Dispose();

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/>
          <Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
          <Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/>
        </Types>
        """;

    private const string DocumentRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="vbaProject.bin" TargetMode="External"/>
        </Relationships>
        """;

    private const string VbaRels = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/>
        </Relationships>
        """;

    private string MakeDocm(bool withMacros)
    {
        string path = _folder.File(withMacros ? "macro.docm" : "plain.docm");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
        Add("[Content_Types].xml", ContentTypesXml);
        Add("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
        Add("word/styles.xml", "<styles/>");
        Add("word/_rels/document.xml.rels", DocumentRels);
        if (withMacros)
        {
            Add("word/vbaProject.bin", "MACRO BYTES");
            Add("word/vbaData.xml", "<data/>");
            Add("word/_rels/vbaProject.bin.rels", VbaRels);
            Add("word/vbaProjectSignature.bin", "SIGNATURE");
        }
        return path;
    }

    private static string Read(ZipArchive zip, string name)
    {
        using var reader = new StreamReader(zip.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }

    [Theory]
    [InlineData("a.docm", true)]
    [InlineData("a.XLSM", true)]
    [InlineData("a.xlsb", true)]
    [InlineData("a.pptm", true)]
    [InlineData("a.docx", false)]
    [InlineData("a.doc", false)]
    public void CanContainMacros_by_extension(string name, bool expected) =>
        Assert.Equal(expected, MacroStripper.CanContainMacros(name));

    [Fact]
    public void Strip_removes_macro_parts_and_everything_that_points_to_them()
    {
        string source = MakeDocm(withMacros: true);
        string copy = _folder.File("copy.docm");

        int removed = MacroStripper.Strip(source, copy);

        Assert.Equal(3, removed);
        using var zip = ZipFile.OpenRead(copy);
        var names = zip.Entries.Select(e => e.FullName).ToList();
        Assert.DoesNotContain(names, n => n.Contains("vba", StringComparison.OrdinalIgnoreCase));
        Assert.Contains("word/document.xml", names);
        Assert.Contains("word/styles.xml", names);

        string types = Read(zip, "[Content_Types].xml");
        Assert.DoesNotContain("vbaProject", types);
        Assert.DoesNotContain("vbaData", types);
        Assert.Contains("/word/document.xml", types);

        string rels = Read(zip, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rId1", rels);
        Assert.Contains("styles.xml", rels);
        Assert.Contains("rId3", rels); // 外部への同名のリンクは部品ではないので残す
    }

    [Fact]
    public void Strip_leaves_the_original_file_untouched()
    {
        string source = MakeDocm(withMacros: true);
        byte[] before = File.ReadAllBytes(source);

        MacroStripper.Strip(source, _folder.File("copy.docm"));

        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Fact]
    public void Strip_copies_a_file_without_macros_as_it_is()
    {
        string source = MakeDocm(withMacros: false);
        string copy = _folder.File("copy.docm");

        Assert.Equal(0, MacroStripper.Strip(source, copy));
        Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(copy));
    }

    [Fact]
    public void Strip_refuses_a_file_that_is_not_a_zip()
    {
        // 暗号化された Office ファイルは zip ではなく OLE 複合ファイルになる
        string source = _folder.File("encrypted.docm");
        File.WriteAllBytes(source, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);
        string copy = _folder.File("copy.docm");

        Assert.Throws<InvalidDataException>(() => MacroStripper.Strip(source, copy));
        Assert.False(File.Exists(copy));
    }
}
