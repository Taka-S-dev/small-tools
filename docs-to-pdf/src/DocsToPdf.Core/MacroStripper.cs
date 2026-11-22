using System.IO.Compression;
using System.Xml.Linq;

namespace DocsToPdf.Core;

/// <summary>
/// マクロ付きの Office ファイル（.docm .xlsm .xlsb .pptm など）の写しから、マクロを取り除く。
///
/// Office はマクロを強制的に無効にして開いているので、マクロが動くことはない。これは二重の守りで、
/// マクロの本体そのものを Office に渡さないようにする。
/// これらの形式は中身が zip なので、マクロの部品（vbaProject.bin と、その署名や付随データ）を除いた zip を書き出し、
/// それらを指している登録（[Content_Types].xml と .rels）も消して、壊れたファイルに見えないようにする。
/// 部品の名前は自由に付けられ、Office は名前ではなく登録の種類で部品をたどるので、名前と登録の両方で探す。
/// 古い形式（.doc .xls .ppt）は構造が違うので対象外で、マクロを無効にする設定だけで守る。
/// </summary>
internal static class MacroStripper
{
    private static readonly HashSet<string> MacroCapable = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".xlsb", ".pptm", ".potm", ".ppsm",
    };

    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>マクロを含められる形式か（拡張子で判断する）。</summary>
    public static bool CanContainMacros(string path) => MacroCapable.Contains(Path.GetExtension(path));

    /// <summary>
    /// source のマクロを除いた写しを destination に書く。
    /// </summary>
    /// <returns>取り除いた部品の数。0 ならマクロは入っていなかった（そのまま写した）。</returns>
    /// <exception cref="InvalidDataException">zip として読めない（暗号化されている、壊れている）。</exception>
    public static int Strip(string source, string destination)
    {
        using var input = ZipFile.OpenRead(source);
        var removed = FindMacroParts(input);
        if (removed.Count == 0)
        {
            input.Dispose();
            File.Copy(source, destination, overwrite: true);
            return 0;
        }

        string temp = destination + ".tmp";
        try
        {
            using (var output = ZipFile.Open(temp, ZipArchiveMode.Create))
            {
                foreach (var entry in input.Entries)
                {
                    if (removed.Contains(entry.FullName) || IsRelationshipsOf(entry.FullName, removed)) continue;

                    var copy = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                    copy.LastWriteTime = entry.LastWriteTime;
                    using var to = copy.Open();
                    using var from = entry.Open();
                    if (entry.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                        WithoutContentTypes(from, removed).Save(to, SaveOptions.DisableFormatting);
                    else if (IsRelationshipsPart(entry.FullName))
                        WithoutRelationships(from, entry.FullName, removed).Save(to, SaveOptions.DisableFormatting);
                    else
                        from.CopyTo(to);
                }
            }
            File.Move(temp, destination, overwrite: true);
            return removed.Count;
        }
        finally
        {
            try { if (File.Exists(temp)) File.Delete(temp); } catch { /* 作業フォルダごと後で消す */ }
        }
    }

    /// <summary>
    /// マクロの本体、その署名、マクロが使うデータの部品を探す。次のどれかに当たれば対象にする:
    /// 決まった名前である / [Content_Types].xml でマクロの種類に登録されている / .rels でマクロとして指されている。
    /// </summary>
    private static HashSet<string> FindMacroParts(ZipArchive zip)
    {
        var entries = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        found.UnionWith(entries.Where(HasMacroName));

        if (zip.GetEntry("[Content_Types].xml") is { } types)
        {
            XElement? root;
            using (var stream = types.Open()) root = XDocument.Load(stream).Root;
            var overridden = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in root?.Elements(ContentTypes + "Override") ?? [])
            {
                if ((string?)item.Attribute("PartName") is not { } part) continue;
                overridden.Add(part.TrimStart('/'));
                if (IsMacroContentType((string?)item.Attribute("ContentType"))) found.Add(part.TrimStart('/'));
            }
            // 拡張子ごとの既定の種類は、名前ごとの登録を持たない部品にだけ効く
            var macroExtensions = (root?.Elements(ContentTypes + "Default") ?? [])
                .Where(d => IsMacroContentType((string?)d.Attribute("ContentType")))
                .Select(d => "." + (string?)d.Attribute("Extension"))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            found.UnionWith(entries.Where(e => !overridden.Contains(e) && macroExtensions.Contains(Path.GetExtension(e))));
        }

        foreach (var rels in zip.Entries.Where(e => IsRelationshipsPart(e.FullName)))
        {
            XElement? root;
            using (var stream = rels.Open()) root = XDocument.Load(stream).Root;
            string baseDir = BaseDirectoryOf(rels.FullName);
            foreach (var r in root?.Elements(Relationships + "Relationship") ?? [])
            {
                if ((string?)r.Attribute("TargetMode") == "External" || (string?)r.Attribute("Target") is not { } target) continue;
                if (IsMacroRelationship((string?)r.Attribute("Type"))) found.Add(Resolve(baseDir, target));
            }
        }

        found.IntersectWith(entries);
        return found;
    }

    /// <summary>
    /// 関係の登録（.rels）か。Office が読むのは _rels フォルダの中のものだけで、
    /// それ以外の場所にある .rels という名前の部品は、ただのデータとしてそのまま写す。
    /// </summary>
    private static bool IsRelationshipsPart(string fullName) =>
        fullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
        && (Path.GetDirectoryName(fullName)?.Replace('\\', '/') ?? "").Split('/')[^1].Equals("_rels", StringComparison.OrdinalIgnoreCase);

    private static bool HasMacroName(string fullName)
    {
        string name = Path.GetFileName(fullName);
        return name.Equals("vbaProject.bin", StringComparison.OrdinalIgnoreCase)
            || name.Equals("vbaData.xml", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("vbaProjectSignature", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>application/vnd.ms-office.vbaProject、vbaProjectSignature…、vnd.ms-word.vbaData+xml など。</summary>
    private static bool IsMacroContentType(string? type) => type is not null
        && (type.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) || type.Contains("vbaData", StringComparison.OrdinalIgnoreCase));

    /// <summary>…/relationships/vbaProject、vbaProjectSignature…、wordVbaData など。</summary>
    private static bool IsMacroRelationship(string? type)
    {
        if (type is null) return false;
        string last = type[(type.LastIndexOf('/') + 1)..];
        return last.StartsWith("vbaProject", StringComparison.OrdinalIgnoreCase) || last.EndsWith("VbaData", StringComparison.OrdinalIgnoreCase);
    }

    private static string BaseDirectoryOf(string relsName) =>
        (SourceOfRelationships(relsName) is { } source ? Path.GetDirectoryName(source)?.Replace('\\', '/') : "") ?? "";

    /// <summary>取り除いた部品自身の .rels（word/_rels/vbaProject.bin.rels など）か。</summary>
    private static bool IsRelationshipsOf(string fullName, HashSet<string> removed)
    {
        if (!fullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) return false;
        string? source = SourceOfRelationships(fullName);
        return source is not null && removed.Contains(source);
    }

    /// <summary>"word/_rels/document.xml.rels" → "word/document.xml"。パッケージ全体の "_rels/.rels" は null。</summary>
    private static string? SourceOfRelationships(string relsName)
    {
        string dir = Path.GetDirectoryName(relsName)?.Replace('\\', '/') ?? "";
        if (!dir.EndsWith("_rels", StringComparison.OrdinalIgnoreCase)) return null;
        string parent = dir[..^"_rels".Length].TrimEnd('/');
        string file = Path.GetFileName(relsName)[..^".rels".Length];
        return file.Length == 0 ? null : (parent.Length > 0 ? parent + "/" : "") + file;
    }

    private static XDocument WithoutContentTypes(Stream stream, HashSet<string> removed)
    {
        var doc = XDocument.Load(stream);
        doc.Root?.Elements(ContentTypes + "Override")
            .Where(o => (string?)o.Attribute("PartName") is { } part && removed.Contains(part.TrimStart('/')))
            .Remove();
        return doc;
    }

    private static XDocument WithoutRelationships(Stream stream, string relsName, HashSet<string> removed)
    {
        var doc = XDocument.Load(stream);
        string baseDir = BaseDirectoryOf(relsName);
        doc.Root?.Elements(Relationships + "Relationship")
            .Where(r => (string?)r.Attribute("TargetMode") != "External"
                        && (string?)r.Attribute("Target") is { } target
                        && removed.Contains(Resolve(baseDir, target)))
            .Remove();
        return doc;
    }

    /// <summary>関係の行き先（相対パス）を、パッケージの中の完全な名前にする。</summary>
    private static string Resolve(string baseDir, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = new List<string>(baseDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var piece in target.Split('/'))
        {
            if (piece == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
            else if (piece is not ("." or "")) parts.Add(piece);
        }
        return string.Join('/', parts);
    }
}
