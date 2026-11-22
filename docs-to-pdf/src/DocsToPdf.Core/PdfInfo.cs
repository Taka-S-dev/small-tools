using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Core;

/// <summary>PDF からページの大きさとしおりを読み取る。</summary>
internal static class PdfInfo
{
    private const int MaxOutlineNodes = 20_000;
    private const int MaxNameTreeNodes = 100_000;

    public static (IReadOnlyList<PageSize> Sizes, IReadOnlyList<OutlineNode> Outline) Read(string path)
    {
        using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return Read(doc);
    }

    public static (IReadOnlyList<PageSize> Sizes, IReadOnlyList<OutlineNode> Outline) Read(PdfDocument doc)
    {
        var sizes = new List<PageSize>(doc.PageCount);
        foreach (var page in doc.Pages)
            sizes.Add(VisibleSize(page));

        IReadOnlyList<OutlineNode> outline;
        try
        {
            outline = ReadOutline(doc);
        }
        catch
        {
            // しおりが壊れていても、結合そのものは続けられる
            outline = [];
        }
        return (sizes, outline);
    }

    /// <summary>ページの見えている範囲（CropBox があればそれ、なければ MediaBox）。</summary>
    public static PdfRectangle VisibleBox(PdfPage page)
    {
        var media = page.MediaBox;
        if (!page.Elements.ContainsKey("/CropBox")) return media;
        var crop = page.CropBox;
        return crop.Width > 0 && crop.Height > 0 ? crop : media;
    }

    public static int Rotation(PdfPage page) => ((page.Rotate % 360) + 360) % 360;

    private static PageSize VisibleSize(PdfPage page)
    {
        var box = VisibleBox(page);
        return Rotation(page) is 90 or 270
            ? new PageSize(box.Height, box.Width)
            : new PageSize(box.Width, box.Height);
    }

    private static List<OutlineNode> ReadOutline(PdfDocument doc)
    {
        var catalog = doc.Internals.Catalog;
        var root = catalog.Elements.GetDictionary("/Outlines");
        if (root is null) return [];

        var pageIndex = new Dictionary<PdfObjectID, int>();
        for (int i = 0; i < doc.PageCount; i++)
        {
            var reference = doc.Pages[i].Reference;
            if (reference is not null) pageIndex[reference.ObjectID] = i;
        }

        var context = new OutlineContext(catalog, pageIndex);
        return ReadLevel(root.Elements.GetDictionary("/First"), context, depth: 0);
    }

    private static List<OutlineNode> ReadLevel(PdfDictionary? item, OutlineContext context, int depth)
    {
        var nodes = new List<OutlineNode>();
        while (item is not null && depth < 32 && context.Visited.Add(item) && context.Visited.Count < MaxOutlineNodes)
        {
            string title = (item.Elements.GetString("/Title") ?? "").Trim();
            int page = context.ResolvePage(item);
            var children = ReadLevel(item.Elements.GetDictionary("/First"), context, depth + 1);

            // 行き先が読めない項目は、子の先頭ページへ飛ばす。子もなければ捨てる
            if (page < 0 && children.Count > 0) page = children[0].PageIndex;
            if (page >= 0 && title.Length > 0)
                nodes.Add(new OutlineNode(title, page, children));
            else
                nodes.AddRange(children);

            item = item.Elements.GetDictionary("/Next");
        }
        return nodes;
    }

    private sealed class OutlineContext(PdfDictionary catalog, Dictionary<PdfObjectID, int> pageIndex)
    {
        private Dictionary<string, PdfItem>? _namedDestinations;

        public HashSet<PdfDictionary> Visited { get; } = new(ReferenceEqualityComparer.Instance);

        public int ResolvePage(PdfDictionary item)
        {
            var dest = Deref(item.Elements["/Dest"]);
            if (dest is null && item.Elements.GetDictionary("/A") is { } action
                && action.Elements.GetName("/S") == "/GoTo")
            {
                dest = Deref(action.Elements["/D"]);
            }
            return ResolveDestination(dest, depth: 0);
        }

        private int ResolveDestination(PdfItem? dest, int depth)
        {
            if (depth > 4) return -1;
            switch (dest)
            {
                case PdfArray array when array.Elements.Count > 0:
                    return array.Elements[0] is PdfReference page && pageIndex.TryGetValue(page.ObjectID, out int index)
                        ? index
                        : -1;
                case PdfDictionary dict:
                    return ResolveDestination(Deref(dict.Elements["/D"]), depth + 1);
                case PdfName name:
                    return ResolveNamed(name.Value.TrimStart('/'), depth);
                case PdfString text:
                    return ResolveNamed(text.Value, depth);
                default:
                    return -1;
            }
        }

        private int ResolveNamed(string name, int depth)
        {
            _namedDestinations ??= LoadNamedDestinations();
            return _namedDestinations.TryGetValue(name, out var dest) ? ResolveDestination(Deref(dest), depth + 1) : -1;
        }

        private Dictionary<string, PdfItem> LoadNamedDestinations()
        {
            var result = new Dictionary<string, PdfItem>(StringComparer.Ordinal);

            // PDF 1.1 の書き方: カタログの /Dests 辞書
            if (catalog.Elements.GetDictionary("/Dests") is { } dests)
            {
                foreach (var key in dests.Elements.Keys)
                {
                    var value = dests.Elements[key];
                    if (value is not null) result[key.TrimStart('/')] = value;
                }
            }

            // PDF 1.2 以降の書き方: /Names の下の名前ツリー
            var tree = catalog.Elements.GetDictionary("/Names")?.Elements.GetDictionary("/Dests");
            if (tree is not null) WalkNameTree(tree, result, new HashSet<PdfDictionary>(ReferenceEqualityComparer.Instance), depth: 0);
            return result;
        }

        /// <summary>
        /// 名前ツリーをたどる。1 つの節を複数の親から指す PDF もあるので、たどった節は二度たどらない
        /// （たどり直すと、深さに対して指数的に時間がかかる）。
        /// </summary>
        private static void WalkNameTree(PdfDictionary node, Dictionary<string, PdfItem> result, HashSet<PdfDictionary> visited, int depth)
        {
            if (depth > 32 || result.Count > 200_000 || visited.Count >= MaxNameTreeNodes || !visited.Add(node)) return;

            if (Deref(node.Elements["/Names"]) is PdfArray names)
            {
                for (int i = 0; i + 1 < names.Elements.Count; i += 2)
                {
                    if (Deref(names.Elements[i]) is PdfString key)
                        result[key.Value] = names.Elements[i + 1];
                }
            }
            if (Deref(node.Elements["/Kids"]) is PdfArray kids)
            {
                foreach (var kid in kids.Elements)
                {
                    if (Deref(kid) is PdfDictionary child)
                        WalkNameTree(child, result, visited, depth + 1);
                }
            }
        }

        private static PdfItem? Deref(PdfItem? item) => item is PdfReference reference ? reference.Value : item;
    }
}
