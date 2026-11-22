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

    /// <summary>マクロの部品に決まった名前を使わず、登録の種類だけでマクロだと分かる docm を作る。</summary>
    private string MakeDocmWithRenamedMacro()
    {
        string path = _folder.File("renamed.docm");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
        Add("[Content_Types].xml", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/>
              <Override PartName="/word/payload.dat" ContentType="application/vnd.ms-office.vbaProject"/>
              <Override PartName="/word/printer.bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings"/>
            </Types>
            """);
        Add("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
        Add("word/_rels/document.xml.rels", """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="payload.dat"/>
              <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="media/code.bin"/>
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="printer.bin"/>
            </Relationships>
            """);
        Add("word/_rels/payload.dat.rels", VbaRels);
        Add("word/payload.dat", "MACRO BYTES");
        Add("word/vbaData.xml", "<data/>");
        Add("word/media/code.bin", "MORE MACRO BYTES");
        Add("word/printer.bin", "PRINTER");
        return path;
    }

    [Fact]
    public void Strip_finds_macro_parts_by_their_registration_even_when_renamed()
    {
        string source = MakeDocmWithRenamedMacro();
        string copy = _folder.File("copy.docm");

        int removed = MacroStripper.Strip(source, copy);

        // payload.dat は名前ごとの登録、code.bin は拡張子ごとの既定、vbaData.xml は名前で見つかる
        Assert.Equal(3, removed);
        using var zip = ZipFile.OpenRead(copy);
        var names = zip.Entries.Select(e => e.FullName).ToList();
        Assert.DoesNotContain("word/payload.dat", names);
        Assert.DoesNotContain("word/_rels/payload.dat.rels", names);
        Assert.DoesNotContain("word/media/code.bin", names);
        Assert.DoesNotContain("word/vbaData.xml", names);
        // 名前ごとに別の種類で登録された .bin（プリンターの設定）は、既定の種類がマクロでも残す
        Assert.Contains("word/printer.bin", names);

        string types = Read(zip, "[Content_Types].xml");
        Assert.DoesNotContain("payload.dat", types);
        Assert.Contains("/word/printer.bin", types);

        string rels = Read(zip, "word/_rels/document.xml.rels");
        Assert.DoesNotContain("rId1", rels);
        Assert.DoesNotContain("rId2", rels);
        Assert.Contains("printer.bin", rels);
    }

    [Fact]
    public void Strip_still_removes_macros_when_a_stray_rels_file_is_not_xml()
    {
        string source = MakeDocm(withMacros: true);
        using (var zip = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            // _rels の外にある .rels は Office が読まない部品。これが壊れていても取り除きは止まらない
            using var writer = new StreamWriter(zip.CreateEntry("word/media/junk.rels").Open());
            writer.Write("<not xml");
        }
        string copy = _folder.File("copy.docm");

        Assert.Equal(3, MacroStripper.Strip(source, copy));
        using var result = ZipFile.OpenRead(copy);
        Assert.DoesNotContain(result.Entries, e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("<not xml", Read(result, "word/media/junk.rels"));
    }

    [Fact]
    public void Strip_refuses_a_package_whose_relationships_are_not_xml()
    {
        string source = MakeDocm(withMacros: true);
        using (var zip = ZipFile.Open(source, ZipArchiveMode.Update))
        {
            zip.GetEntry("word/_rels/document.xml.rels")!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("word/_rels/document.xml.rels").Open());
            writer.Write("<not xml");
        }

        // 呼び出し側（ConversionService）は、zip なのにこれが出たら変換しない
        Assert.ThrowsAny<System.Xml.XmlException>(() => MacroStripper.Strip(source, _folder.File("copy.docm")));
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
