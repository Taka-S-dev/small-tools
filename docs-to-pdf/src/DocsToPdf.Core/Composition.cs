using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocsToPdf.Core;

/// <summary>
/// 一覧の 1 行のうち、利用者が決めたもの。Path と Root は完全なパス。
/// OutlineChanges と PageBookmarks は決めたときのしおりとページの並びを持つので、元のファイルが変わっていれば当てない。
/// </summary>
internal sealed record CompositionItem(string Path, string? Root, string Title)
{
    public string PageRange { get; init; } = "";
    public IReadOnlyList<string>? Sheets { get; init; }
    public bool FitWidth { get; init; }
    public IReadOnlyDictionary<int, int> Rotations { get; init; } = FileEntry.NoRotations;
    public int? OutlineDepth { get; init; }
    public OutlineChanges? OutlineChanges { get; init; }
    public PageBookmarks? PageBookmarks { get; init; }

    public static CompositionItem Of(FileEntry e) => new(e.Path, e.Root, e.Title)
    {
        PageRange = e.PageRange,
        Sheets = e.Sheets,
        FitWidth = e.FitWidth,
        Rotations = e.Rotations,
        OutlineDepth = e.OutlineDepth,
        OutlineChanges = e.OutlineChanges,
        PageBookmarks = e.PageBookmarks,
    };

    /// <summary>この行を、変換した結果に当てて結合する 1 ファイル分にする（フォルダの段は Composition.FoldersOf で付ける）。</summary>
    public MergePart PartOf(ConvertedPdf pdf)
    {
        bool outlineApplies = OutlineChanges?.AppliesTo(pdf.Outline) == true;
        return new MergePart(pdf.PdfPath, Title.Trim() is { Length: > 0 } t ? t : System.IO.Path.GetFileNameWithoutExtension(Path),
            Core.PageRange.Resolve(PageRange, pdf.PageCount), pdf.Outline, Rotations)
        {
            OutlineDepth = OutlineDepth,
            OutlineEdits = outlineApplies ? OutlineChanges!.Edits : MergePart.NoEdits,
            OutlineLevels = outlineApplies ? OutlineChanges!.Levels : MergePart.NoLevels,
            Added = PageBookmarks is { } added && added.AppliesTo(pdf.PageSizes) ? added.Items : [],
        };
    }
}

/// <summary>
/// 一覧の保存（.docstopdf）。並び・行ごとの決めごと・オプションを持ち、開けば同じ PDF を作り直せる。
/// パスワードは保存しない。
/// </summary>
/// <param name="OutlineDepth">元のファイルのしおりを何段まで付けるか。null ならすべて、0 なら付けない。</param>
/// <param name="FolderNames">付け直したフォルダの段の名前（フォルダの完全なパス → 名前）。</param>
internal sealed record Composition(
    IReadOnlyList<CompositionItem> Items,
    bool PageNumbers,
    bool TableOfContents,
    bool DuplexBlanks,
    int? OutlineDepth,
    bool FolderBookmarks,
    IReadOnlyDictionary<string, string> FolderNames)
{
    public const string Extension = ".docstopdf";

    /// <summary>しおりの一番上の段ごとに中表紙を入れるか（MergeOptions.Dividers）。</summary>
    public bool Dividers { get; init; }

    public static bool IsCompositionFile(string path) =>
        string.Equals(System.IO.Path.GetExtension(path), Extension, StringComparison.OrdinalIgnoreCase);

    /// <summary>行ごとの、しおりのフォルダの段（付け直した名前を当てたもの）。フォルダごとにまとめないなら、どれも空。</summary>
    public static IReadOnlyList<IReadOnlyList<BookmarkFolder>> FoldersOf(IReadOnlyList<CompositionItem> items, bool folderBookmarks,
        IReadOnlyDictionary<string, string> folderNames)
    {
        if (!folderBookmarks) return items.Select(_ => (IReadOnlyList<BookmarkFolder>)[]).ToList();
        return BookmarkFolders.Of(items.Select(i => i.Path).ToList(), items.Select(i => i.Root).ToList())
            .Select(folders => (IReadOnlyList<BookmarkFolder>)folders
                .Select(f => folderNames.TryGetValue(f.Key, out var name) ? f with { Name = name } : f).ToList())
            .ToList();
    }

    public MergeOptions MergeOptions(string documentTitle) =>
        new(PageNumbers, TableOfContents, DuplexBlanks, documentTitle) { OutlineDepth = OutlineDepth, Dividers = Dividers };
}

/// <summary>
/// 一覧が、一覧の置き場所とは別のコンピューターの共有フォルダを指している。
/// そこへファイルを探しに行くだけで、Windows はそのコンピューターへ利用者の認証情報を送るので、確かめずには行かない。
/// </summary>
internal sealed class NetworkPathsException(IReadOnlyList<string> hosts)
    : Exception($"一覧が、ほかのコンピューターの共有フォルダ（{string.Join("、", hosts.Select(h => @"\\" + h))}）を指しています。")
{
    public IReadOnlyList<string> Hosts { get; } = hosts;
}

/// <summary>一覧の保存の読み書き。パスは保存したファイルからの相対パスで持ち、フォルダごと移しても開けるようにする。</summary>
internal static class CompositionFile
{
    private const int Version = 1;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static void Save(string path, Composition composition)
    {
        string baseDir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path))!;
        var dto = new FileDto
        {
            Version = Version,
            PageNumbers = composition.PageNumbers,
            TableOfContents = composition.TableOfContents,
            DuplexBlanks = composition.DuplexBlanks,
            Dividers = composition.Dividers ? true : null,
            BookmarkDepth = composition.OutlineDepth,
            FolderBookmarks = composition.FolderBookmarks,
            FolderNames = composition.FolderNames.Count == 0 ? null
                : composition.FolderNames.Select(f => new FolderNameDto { Folder = PathDto.Of(f.Key, baseDir), Name = f.Value }).ToList(),
            Files = composition.Items.Select(i => new ItemDto
            {
                File = PathDto.Of(i.Path, baseDir),
                Root = i.Root is null ? null : PathDto.Of(i.Root, baseDir),
                Title = i.Title,
                Pages = i.PageRange.Length == 0 ? null : i.PageRange,
                Sheets = i.Sheets?.ToList(),
                FitWidth = i.FitWidth ? true : null,
                Rotations = i.Rotations.Count == 0 ? null : i.Rotations.ToDictionary(r => r.Key, r => r.Value),
                BookmarkDepth = i.OutlineDepth,
                OutlineChanges = i.OutlineChanges is { } c
                    ? new OutlineChangesDto
                    {
                        Outline = c.Outline.Select(NodeDto.Of).ToList(),
                        Renames = c.Edits.Count == 0 ? null : c.Edits.ToDictionary(e => e.Key, e => e.Value),
                        Levels = c.Levels.Count == 0 ? null : c.Levels.ToDictionary(l => l.Key, l => l.Value),
                    }
                    : null,
                AddedBookmarks = i.PageBookmarks is { } b
                    ? new AddedBookmarksDto
                    {
                        PageSizes = b.Pages.Select(p => new[] { p.Width, p.Height }).ToList(),
                        Items = b.Items.Select(a => new AddedDto { Title = a.Title, Page = a.Page, Level = a.Level }).ToList(),
                    }
                    : null,
            }).ToList(),
        };

        // 書きかけのファイルを残さないよう、隣に書いてから置き換える
        string temp = path + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(dto, Json));
        File.Move(temp, path, overwrite: true);
    }

    /// <param name="allowNetwork">
    /// 一覧が、一覧の置き場所とは別のコンピューターの共有フォルダを指しているときに、そのコンピューターの名前を渡して、使ってよいかを尋ねる。
    /// null か false なら NetworkPathsException を投げ、そこへは探しに行かない。
    /// </param>
    /// <exception cref="FormatException">一覧の保存として読めない。</exception>
    /// <exception cref="IOException">ファイルを読めない。</exception>
    /// <exception cref="NetworkPathsException">ほかのコンピューターの共有フォルダを指していて、使ってよいとされなかった。</exception>
    public static Composition Load(string path, Func<IReadOnlyList<string>, bool>? allowNetwork = null)
    {
        string fullPath = System.IO.Path.GetFullPath(path);
        string baseDir = System.IO.Path.GetDirectoryName(fullPath)!;
        FileDto? dto;
        try
        {
            dto = JsonSerializer.Deserialize<FileDto>(File.ReadAllText(path), Json);
        }
        catch (JsonException ex)
        {
            throw new FormatException("一覧の保存として読めません（ファイルが壊れているか、別の種類のファイルです）。", ex);
        }
        if (dto?.Files is null) throw new FormatException("一覧の保存として読めません（ファイルの一覧がありません）。");
        if (dto.Version > Version) throw new FormatException("新しい版のアプリで保存された一覧です。アプリを新しくしてから開いてください。");

        // ほかのコンピューターへ探しに行く前に確かめる（一覧と同じコンピューターの共有フォルダなら、一覧を開いた時点でつながっている）
        var paths = dto.Files.SelectMany(f => new[] { (f.File, true), (f.Root, false) })
            .Concat((dto.FolderNames ?? []).Select(f => (f.Folder, false)));
        string? ownHost = HostOf(fullPath);
        var hosts = paths.Where(p => p.Item1 is not null)
            .Select(p => p.Item1!.ForeignHost(baseDir, p.Item2, ownHost))
            .OfType<string>().Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (hosts.Count > 0 && allowNetwork?.Invoke(hosts) != true) throw new NetworkPathsException(hosts);

        try
        {
            var items = dto.Files.Select(f =>
            {
                if (f.File is null) throw new FormatException("ファイルのパスがない行があります。");
                string file = f.File.Resolve(baseDir, expectFile: true);
                return new CompositionItem(file, f.Root?.Resolve(baseDir, expectFile: false), f.Title ?? System.IO.Path.GetFileNameWithoutExtension(file))
                {
                    PageRange = f.Pages is { } pages ? PageRange.Normalize(pages) : "",
                    Sheets = f.Sheets,
                    FitWidth = f.FitWidth == true,
                    Rotations = f.Rotations is { Count: > 0 } r
                        ? r.Where(x => x.Key >= 0 && x.Value is 90 or 180 or 270).ToDictionary(x => x.Key, x => x.Value)
                        : FileEntry.NoRotations,
                    OutlineDepth = f.BookmarkDepth is < 0 ? null : f.BookmarkDepth,
                    OutlineChanges = f.OutlineChanges is { Outline: { } outline } c
                        ? new OutlineChanges(outline.Select(n => n.ToNode()).ToList(),
                            c.Renames ?? new Dictionary<string, string?>())
                        {
                            Levels = c.Levels ?? new Dictionary<string, int>(),
                        }
                        : null,
                    PageBookmarks = f.AddedBookmarks is { PageSizes: { } sizes, Items: { } added }
                        ? new PageBookmarks(
                            sizes.Select(s => s is [var w, var h] ? new PageSize(w, h) : throw new FormatException("ページの大きさが読めません。")).ToList(),
                            added.Select(a => new AddedBookmark(a.Title ?? "", a.Page, Math.Max(1, a.Level))).ToList())
                        : null,
                };
            }).ToList();

            var folderNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in dto.FolderNames ?? [])
                if (f.Folder is not null && f.Name is not null) folderNames[f.Folder.Resolve(baseDir, expectFile: false)] = f.Name;

            return new Composition(items, dto.PageNumbers, dto.TableOfContents, dto.DuplexBlanks,
                dto.BookmarkDepth is < 0 ? null : dto.BookmarkDepth, dto.FolderBookmarks, folderNames)
            {
                Dividers = dto.Dividers == true,
            };
        }
        catch (ArgumentException ex)
        {
            throw new FormatException("一覧の保存に、パスとして読めない値があります。", ex);
        }
    }

    // ── ファイルの形 ──
    // 名前は保存したファイルの中身そのものなので、変えると前に保存したものが開けなくなる

    private sealed class FileDto
    {
        public int Version { get; set; }
        public bool PageNumbers { get; set; }
        public bool TableOfContents { get; set; }
        public bool DuplexBlanks { get; set; }
        public bool? Dividers { get; set; }
        public int? BookmarkDepth { get; set; }
        public bool FolderBookmarks { get; set; } = true;
        public List<FolderNameDto>? FolderNames { get; set; }
        public List<ItemDto>? Files { get; set; }
    }

    private sealed class ItemDto
    {
        public PathDto? File { get; set; }
        public PathDto? Root { get; set; }
        public string? Title { get; set; }
        public string? Pages { get; set; }
        public List<string>? Sheets { get; set; }
        public bool? FitWidth { get; set; }
        public Dictionary<int, int>? Rotations { get; set; }
        public int? BookmarkDepth { get; set; }
        public OutlineChangesDto? OutlineChanges { get; set; }
        public AddedBookmarksDto? AddedBookmarks { get; set; }
    }

    /// <summary>
    /// パスは相対パスと絶対パスの両方を持つ。保存したファイルと一緒に移したなら相対パスで、
    /// 保存したファイルだけを移したなら絶対パスで見つける。
    /// </summary>
    private sealed class PathDto
    {
        public string? Relative { get; set; }
        public string? Absolute { get; set; }

        public static PathDto Of(string fullPath, string baseDir)
        {
            string relative = System.IO.Path.GetRelativePath(baseDir, fullPath);
            // 別のドライブなら相対にできない（GetRelativePath がそのまま返す）
            return new PathDto { Relative = System.IO.Path.IsPathRooted(relative) ? null : relative, Absolute = fullPath };
        }

        /// <summary>
        /// 一覧は人から受け取ることもあるので、ふつうのファイルを指さない書き方は読まない。
        /// 相対パスの欄に書かれた絶対パスと、デバイスのパス（\\?\ や \\.\）は使わない。
        /// ほかのコンピューターの共有フォルダは、Load で使ってよいと確かめてから探しに行く。
        /// </summary>
        public string Resolve(string baseDir, bool expectFile)
        {
            var (relative, absolute) = Candidates(baseDir);
            if (relative is not null && Exists(relative, expectFile)) return relative;
            if (absolute is not null && Exists(absolute, expectFile)) return absolute;
            return relative ?? absolute ?? throw new FormatException("パスのない項目か、使えないパスの項目があります。");
        }

        /// <summary>
        /// Resolve が、一覧の置き場所とは別のコンピューター（ownHost 以外）の共有フォルダへ探しに行くなら、そのコンピューターの名前。
        /// 相対パスで見つかるなら、絶対パスの側へは行かないので null。
        /// </summary>
        public string? ForeignHost(string baseDir, bool expectFile, string? ownHost)
        {
            var (relative, absolute) = Candidates(baseDir);
            if (relative is not null && Exists(relative, expectFile)) return null;
            return absolute is not null && HostOf(absolute) is { } host && !string.Equals(host, ownHost, StringComparison.OrdinalIgnoreCase)
                ? host : null;
        }

        private (string? Relative, string? Absolute) Candidates(string baseDir)
        {
            string? relative = Relative is { Length: > 0 } r && !System.IO.Path.IsPathRooted(r)
                ? System.IO.Path.GetFullPath(System.IO.Path.Combine(baseDir, r)) : null;
            string? absolute = Absolute is { Length: > 0 } a && System.IO.Path.IsPathFullyQualified(a) && !IsDevicePath(a)
                ? System.IO.Path.GetFullPath(a) : null;
            // 書き方を整えたあとでも、デバイスのパスにはしない
            return (relative is not null && IsDevicePath(relative) ? null : relative, absolute is not null && IsDevicePath(absolute) ? null : absolute);
        }

        private static bool Exists(string path, bool expectFile) => expectFile ? System.IO.File.Exists(path) : Directory.Exists(path);
    }

    /// <summary>区切りを \ にそろえる（"/" と "\" を混ぜた書き方で、判定をすり抜けさせない）。</summary>
    private static string Separators(string path) => path.Replace('/', '\\');

    /// <summary>デバイスのパス（\\?\、\\.\、\??\）か。</summary>
    private static bool IsDevicePath(string path)
    {
        string p = Separators(path);
        return p.StartsWith(@"\\?\", StringComparison.Ordinal) || p.StartsWith(@"\\.\", StringComparison.Ordinal)
            || p.StartsWith(@"\??\", StringComparison.Ordinal);
    }

    /// <summary>
    /// 共有フォルダのパス（\\サーバー\共有\...）なら、そのコンピューターの名前。ほかは null。
    /// 名前が読めないときは "?" を返し、ほかのコンピューターとして扱わせる（確かめずに探しに行かないように）。
    /// </summary>
    internal static string? HostOf(string path)
    {
        string p = Separators(path);
        if (!p.StartsWith(@"\\", StringComparison.Ordinal) || IsDevicePath(p)) return null;
        string rest = p[2..];
        int end = rest.IndexOf('\\');
        string host = end < 0 ? rest : rest[..end];
        return host.Length > 0 ? host : "?";
    }

    private sealed class FolderNameDto
    {
        public PathDto? Folder { get; set; }
        public string? Name { get; set; }
    }

    private sealed class OutlineChangesDto
    {
        public List<NodeDto>? Outline { get; set; }
        public Dictionary<string, string?>? Renames { get; set; }
        public Dictionary<string, int>? Levels { get; set; }
    }

    private sealed class NodeDto
    {
        public string Title { get; set; } = "";
        public int Page { get; set; }
        public List<NodeDto>? Children { get; set; }

        public static NodeDto Of(OutlineNode n) => new()
        {
            Title = n.Title,
            Page = n.PageIndex,
            Children = n.Children.Count == 0 ? null : n.Children.Select(Of).ToList(),
        };

        public OutlineNode ToNode() => new(Title, Page, Children?.Select(c => c.ToNode()).ToList() ?? []);
    }

    private sealed class AddedBookmarksDto
    {
        public List<double[]>? PageSizes { get; set; }
        public List<AddedDto>? Items { get; set; }
    }

    private sealed class AddedDto
    {
        public string? Title { get; set; }
        public int Page { get; set; }
        public int Level { get; set; } = 1;
    }
}
