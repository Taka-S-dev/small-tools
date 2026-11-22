using DocsToPdf.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocsToPdf.Tests;

public class NamedDestinationTests
{
    /// <summary>
    /// しおりの行き先を名前（"target"）で書き、その名前を /Names の名前ツリーに置いた PDF を作る。
    /// 名前ツリーは、1 つの子を 2 回ずつ指す節を levels 段重ねる。
    /// </summary>
    private static string PdfWithSharedNameTree(TestFolder folder, int levels)
    {
        string path = folder.File("named.pdf");
        using (var doc = new PdfDocument())
        {
            doc.AddPage();
            var target = doc.AddPage();

            var leaf = new PdfDictionary(doc);
            var destination = new PdfArray(doc);
            destination.Elements.Add(target.Reference!);
            destination.Elements.Add(new PdfName("/Fit"));
            var names = new PdfArray(doc);
            names.Elements.Add(new PdfString("target"));
            names.Elements.Add(destination);
            leaf.Elements["/Names"] = names;
            doc.Internals.AddObject(leaf);

            var node = leaf;
            for (int i = 0; i < levels; i++)
            {
                var parent = new PdfDictionary(doc);
                var kids = new PdfArray(doc);
                kids.Elements.Add(node.Reference!);
                kids.Elements.Add(node.Reference!);
                parent.Elements["/Kids"] = kids;
                doc.Internals.AddObject(parent);
                node = parent;
            }

            var nameDictionary = new PdfDictionary(doc);
            nameDictionary.Elements["/Dests"] = node.Reference!;
            doc.Internals.Catalog.Elements["/Names"] = nameDictionary;

            // PDFsharp のしおりは保存のときに行き先のページを直接書くので、名前で引くしおりは辞書を自分で組む
            var root = new PdfDictionary(doc);
            var item = new PdfDictionary(doc);
            doc.Internals.AddObject(root);
            doc.Internals.AddObject(item);
            var goTo = new PdfDictionary(doc);
            goTo.Elements["/S"] = new PdfName("/GoTo");
            goTo.Elements["/D"] = new PdfString("target");
            item.Elements["/Title"] = new PdfString("章");
            item.Elements["/Parent"] = root.Reference!;
            item.Elements["/A"] = goTo;
            root.Elements["/Type"] = new PdfName("/Outlines");
            root.Elements["/First"] = item.Reference!;
            root.Elements["/Last"] = item.Reference!;
            root.Elements["/Count"] = new PdfInteger(1);
            doc.Internals.Catalog.Elements["/Outlines"] = root.Reference!;
            doc.Save(path);
        }
        return path;
    }

    [Fact]
    public void Resolves_a_named_destination_through_the_name_tree()
    {
        using var folder = new TestFolder();
        string path = PdfWithSharedNameTree(folder, levels: 1);

        var outline = PdfInfo.Read(path).Outline;

        Assert.Equal(1, Assert.Single(outline).PageIndex);
    }

    [Fact]
    public async Task Reads_a_name_tree_whose_nodes_are_shared_without_walking_them_again()
    {
        using var folder = new TestFolder();
        // たどり直すと 2^30 回になり、終わらない
        string path = PdfWithSharedNameTree(folder, levels: 30);
        var ct = TestContext.Current.CancellationToken;

        var outline = await Task.Run(() => PdfInfo.Read(path).Outline, ct).WaitAsync(TimeSpan.FromSeconds(10), ct);

        Assert.Equal(1, Assert.Single(outline).PageIndex);
    }
}
