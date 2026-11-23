using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public class CompositionTests
{
    private static readonly IReadOnlyList<OutlineNode> Outline =
        [new("第1章", 0, [new("1.1", 1, [])]), new("第2章", 2, [])];

    private static readonly IReadOnlyList<PageSize> Sizes = [new(595.28, 841.89), new(595.28, 841.89), new(841.89, 595.28)];

    private static Composition Sample(string a, string b, string root) => new(
        [
            new CompositionItem(a, root, "報告書")
            {
                PageRange = "1-2",
                Sheets = ["集計", "明細"],
                FitWidth = true,
                Rotations = new Dictionary<int, int> { [1] = 90 },
                OutlineDepth = 2,
                OutlineChanges = new OutlineChanges(Outline, new Dictionary<string, string?> { ["0"] = "はじめに", ["1"] = null })
                {
                    Levels = new Dictionary<string, int> { ["0/0"] = 1 },
                },
                PageBookmarks = new PageBookmarks(Sizes, [new AddedBookmark("付録", 2, 2)]),
            },
            new CompositionItem(b, null, "b"),
        ],
        PageNumbers: true, TableOfContents: false, DuplexBlanks: true, OutlineDepth: 1, FolderBookmarks: false,
        new Dictionary<string, string> { [root] = "資料" })
    {
        Dividers = true,
    };

    [Fact]
    public void Save_then_Load_keeps_every_choice()
    {
        using var folder = new TestFolder();
        string root = Directory.CreateDirectory(folder.File("docs")).FullName;
        string a = folder.Pdf(Path.Combine("docs", "a.pdf"), 3);
        string b = folder.Pdf("b.pdf", 1);
        string saved = folder.File("list" + Composition.Extension);

        CompositionFile.Save(saved, Sample(a, b, root));
        var loaded = CompositionFile.Load(saved);

        Assert.True(loaded.PageNumbers);
        Assert.False(loaded.TableOfContents);
        Assert.True(loaded.DuplexBlanks);
        Assert.Equal(1, loaded.OutlineDepth);
        Assert.False(loaded.FolderBookmarks);
        Assert.True(loaded.Dividers);
        Assert.Equal("資料", loaded.FolderNames[root]);
        Assert.Equal(2, loaded.Items.Count);

        var item = loaded.Items[0];
        Assert.Equal(a, item.Path);
        Assert.Equal(root, item.Root);
        Assert.Equal("報告書", item.Title);
        Assert.Equal("1-2", item.PageRange);
        Assert.Equal(["集計", "明細"], item.Sheets!);
        Assert.True(item.FitWidth);
        Assert.Equal(90, item.Rotations[1]);
        Assert.Equal(2, item.OutlineDepth);
        Assert.True(item.OutlineChanges!.AppliesTo(Outline));
        Assert.Equal("はじめに", item.OutlineChanges.Edits["0"]);
        Assert.Null(item.OutlineChanges.Edits["1"]);
        Assert.Equal(1, item.OutlineChanges.Levels["0/0"]);
        Assert.True(item.PageBookmarks!.AppliesTo(Sizes));
        Assert.Equal(new AddedBookmark("付録", 2, 2), item.PageBookmarks.Items.Single());

        var plain = loaded.Items[1];
        Assert.Equal(b, plain.Path);
        Assert.Null(plain.Root);
        Assert.Equal("", plain.PageRange);
        Assert.Null(plain.Sheets);
        Assert.Empty(plain.Rotations);
        Assert.Null(plain.OutlineDepth);
        Assert.Null(plain.OutlineChanges);
        Assert.Null(plain.PageBookmarks);
    }

    [Fact]
    public void Load_finds_files_moved_together_with_the_saved_list()
    {
        using var folder = new TestFolder();
        string a = folder.Pdf("a.pdf", 1);
        CompositionFile.Save(folder.File("list" + Composition.Extension), new Composition(
            [new CompositionItem(a, null, "a")], false, false, false, null, true, new Dictionary<string, string>()));

        // フォルダごと別の場所へ移す
        using var other = new TestFolder();
        string moved = Path.Combine(other.Path, "moved");
        Directory.Move(folder.Path, moved);
        try
        {
            var loaded = CompositionFile.Load(Path.Combine(moved, "list" + Composition.Extension));
            Assert.Equal(Path.Combine(moved, "a.pdf"), loaded.Items[0].Path);
        }
        finally
        {
            Directory.Move(moved, folder.Path);
        }
    }

    [Fact]
    public void Load_falls_back_to_the_absolute_path_when_only_the_list_was_moved()
    {
        using var folder = new TestFolder();
        string a = folder.Pdf("a.pdf", 1);
        string saved = folder.File("list" + Composition.Extension);
        CompositionFile.Save(saved, new Composition(
            [new CompositionItem(a, null, "a")], false, false, false, null, true, new Dictionary<string, string>()));

        using var other = new TestFolder();
        string moved = Path.Combine(other.Path, "list" + Composition.Extension);
        File.Move(saved, moved);

        Assert.Equal(a, CompositionFile.Load(moved).Items[0].Path);
    }

    [Theory]
    [InlineData(@"\\?\C:\a.pdf")]
    [InlineData(@"\\.\C:\a.pdf")]
    [InlineData("C:a.pdf")]
    [InlineData(@"\??\C:\a.pdf")]
    [InlineData(@"\/?/C:/a.pdf")]
    [InlineData("//?/C:/a.pdf")]
    public void Load_refuses_device_and_drive_relative_paths(string absolute)
    {
        using var folder = new TestFolder();
        string path = folder.File("bad" + Composition.Extension);
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new { version = 1, files = new[] { new { file = new { absolute } } } }));
        Assert.Throws<FormatException>(() => CompositionFile.Load(path));
    }

    [Theory]
    [InlineData(@"\\server\share\a.pdf", "server")]
    [InlineData("//server/share/a.pdf", "server")]
    [InlineData(@"/\server/share/a.pdf", "server")]
    [InlineData(@"\\\share\a.pdf", "?")]
    [InlineData(@"C:\docs\a.pdf", null)]
    [InlineData(@"\\?\UNC\server\share\a.pdf", null)]
    public void HostOf_names_the_computer_of_a_shared_folder(string path, string? host) =>
        Assert.Equal(host, CompositionFile.HostOf(path));

    [Fact]
    public void Load_does_not_reach_another_computer_without_asking()
    {
        using var folder = new TestFolder();
        string path = folder.File("list" + Composition.Extension);
        string json = System.Text.Json.JsonSerializer.Serialize(new
        {
            version = 1,
            files = new[] { new { file = new { relative = "missing.pdf", absolute = @"\\other-pc\share\a.pdf" } } },
        });
        File.WriteAllText(path, json);

        Assert.Equal(["other-pc"], Assert.Throws<NetworkPathsException>(() => CompositionFile.Load(path)).Hosts);
        IReadOnlyList<string>? asked = null;
        Assert.Throws<NetworkPathsException>(() => CompositionFile.Load(path, hosts => { asked = hosts; return false; }));
        Assert.Equal(["other-pc"], asked!);
    }

    [Fact]
    public void Load_does_not_ask_when_the_relative_path_is_found()
    {
        using var folder = new TestFolder();
        folder.Pdf("a.pdf", 1);
        string path = folder.File("list" + Composition.Extension);
        string json = System.Text.Json.JsonSerializer.Serialize(new
        {
            version = 1,
            files = new[] { new { file = new { relative = "a.pdf", absolute = @"\\other-pc\share\a.pdf" } } },
        });
        File.WriteAllText(path, json);

        var loaded = CompositionFile.Load(path, _ => throw new InvalidOperationException("尋ねてはいけない"));
        Assert.Equal(folder.File("a.pdf"), loaded.Items[0].Path);
    }

    [Fact]
    public void Load_ignores_an_absolute_path_written_as_the_relative_one()
    {
        using var folder = new TestFolder();
        string a = folder.Pdf("a.pdf", 1);
        string path = folder.File("list" + Composition.Extension);
        string json = System.Text.Json.JsonSerializer.Serialize(new { version = 1, files = new[] { new { file = new { relative = a } } } });
        File.WriteAllText(path, json);
        Assert.Throws<FormatException>(() => CompositionFile.Load(path));
    }

    [Theory]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("{\"version\": 99, \"files\": []}")]
    public void Load_rejects_what_is_not_a_saved_list(string content)
    {
        using var folder = new TestFolder();
        string path = folder.File("bad" + Composition.Extension);
        File.WriteAllText(path, content);
        Assert.Throws<FormatException>(() => CompositionFile.Load(path));
    }

    [Fact]
    public void PartOf_applies_the_bookmark_edits_only_to_the_same_outline_and_pages()
    {
        var item = Sample("C:\\a.docx", "C:\\b.pdf", "C:\\docs").Items[0];

        var same = item.PartOf(new ConvertedPdf("a.pdf", Sizes, Outline, null));
        Assert.Equal("報告書", same.Title);
        Assert.Equal([0, 1], same.Pages);
        Assert.Equal(90, same.TurnOf(1));
        Assert.Equal(2, same.OutlineDepth);
        Assert.Equal("はじめに", same.OutlineEdits["0"]);
        Assert.Equal(1, same.OutlineLevels["0/0"]);
        Assert.Single(same.Added);

        // 元のファイルが書き換わって、しおりもページの並びも変わった
        IReadOnlyList<OutlineNode> changedOutline = [new("序章", 0, [])];
        IReadOnlyList<PageSize> changedSizes = [new(595.28, 841.89)];
        var changed = item.PartOf(new ConvertedPdf("a.pdf", changedSizes, changedOutline, null));
        Assert.Empty(changed.OutlineEdits);
        Assert.Empty(changed.OutlineLevels);
        Assert.Empty(changed.Added);
        Assert.Equal([0], changed.Pages);
    }

    [Fact]
    public void PartOf_uses_the_file_name_for_an_empty_title()
    {
        var part = new CompositionItem("C:\\docs\\報告.docx", null, "  ").PartOf(new ConvertedPdf("x.pdf", Sizes, [], null));
        Assert.Equal("報告", part.Title);
    }

    [Fact]
    public void FoldersOf_applies_renamed_folders()
    {
        string root = Path.GetFullPath("C:\\docs");
        CompositionItem[] items = [new(Path.Combine(root, "sub", "a.pdf"), root, "a")];
        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { [Path.Combine(root, "sub")] = "別名" };

        Assert.Equal("別名", Composition.FoldersOf(items, true, names)[0].Single().Name);
        Assert.Empty(Composition.FoldersOf(items, false, names)[0]);
    }
}
