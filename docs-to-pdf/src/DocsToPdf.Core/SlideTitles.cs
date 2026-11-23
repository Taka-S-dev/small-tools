using System.Text.RegularExpressions;

namespace DocsToPdf.Core;

/// <summary>
/// PowerPoint が PDF に付けるしおりは「スライド 3: まとめ」の形で、番号はページを見れば分かるので外す。
/// 英語版の Office では「Slide 3: まとめ」になる。タイトルのないスライドは「スライド 3」だけなので、そのまま残す。
/// </summary>
internal static partial class SlideTitles
{
    [GeneratedRegex(@"^(?:スライド|Slide)\s*\d+\s*[:：]\s*(?<title>\S.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex Numbered();

    public static string Strip(string title) => Numbered().Match(title) is { Success: true } m ? m.Groups["title"].Value : title;

    public static IReadOnlyList<OutlineNode> Strip(IReadOnlyList<OutlineNode> nodes) =>
        nodes.Select(n => n with { Title = Strip(n.Title), Children = Strip(n.Children) }).ToList();
}
