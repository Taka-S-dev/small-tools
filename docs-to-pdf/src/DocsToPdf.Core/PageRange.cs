using System.Text;

namespace DocsToPdf.Core;

/// <summary>
/// 「1-3,5,8-」のようなページ指定。番号は 1 始まり。「8-」は 8 ページ以降、「-3」は 3 ページまで。
/// 結果はいつもページ順（重複なし）で、書いた順には並べ替えない。
/// </summary>
internal static class PageRange
{
    /// <summary>全角の数字や記号、空白を取り除いて、半角の書き方にそろえる。</summary>
    public static string Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        var sb = new StringBuilder(text.Length);
        foreach (char ch in text)
        {
            char c = ch switch
            {
                >= '０' and <= '９' => (char)('0' + (ch - '０')),
                '，' or '、' or ';' or '；' or '・' => ',',
                '－' or '‐' or '–' or '—' or '―' or 'ー' or '〜' or '～' or '~' => '-',
                _ => ch,
            };
            if (!char.IsWhiteSpace(c)) sb.Append(c);
        }
        return sb.ToString().Trim(',');
    }

    public static bool IsValid(string text) => TryParse(Normalize(text), out _);

    private static bool TryParse(string normalized, out List<(int From, int? To)> ranges)
    {
        ranges = [];
        if (normalized.Length == 0) return true;

        foreach (var part in normalized.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            int dash = part.IndexOf('-');
            if (dash < 0)
            {
                if (!TryNumber(part, out int page)) return false;
                ranges.Add((page, page));
                continue;
            }
            if (part.IndexOf('-', dash + 1) >= 0) return false;

            string left = part[..dash], right = part[(dash + 1)..];
            if (left.Length == 0 && right.Length == 0) return false;
            int from = 1;
            int? to = null;
            if (left.Length > 0 && !TryNumber(left, out from)) return false;
            if (right.Length > 0)
            {
                if (!TryNumber(right, out int end)) return false;
                to = end;
            }
            if (to is int t && t < from) return false;
            ranges.Add((from, to));
        }
        return true;
    }

    private static bool TryNumber(string text, out int value) =>
        int.TryParse(text, System.Globalization.NumberStyles.None, null, out value) && value >= 1;

    /// <summary>使うページを 0 始まりの番号で返す。指定が空か読めなければすべて、範囲外の番号は無視する。</summary>
    public static IReadOnlyList<int> Resolve(string text, int pageCount)
    {
        string normalized = Normalize(text);
        if (normalized.Length == 0 || !TryParse(normalized, out var ranges))
            return Enumerable.Range(0, pageCount).ToArray();

        var included = new bool[pageCount];
        foreach (var (from, to) in ranges)
        {
            int last = Math.Min(to ?? pageCount, pageCount);
            for (int page = from; page <= last; page++)
                included[page - 1] = true;
        }
        var result = new List<int>();
        for (int i = 0; i < pageCount; i++)
            if (included[i]) result.Add(i);
        return result;
    }

    /// <summary>0 始まりのページ番号の集まりを、短い指定に直す。全ページなら空文字。</summary>
    public static string Format(IEnumerable<int> pages, int pageCount)
    {
        var sorted = pages.Distinct().Order().ToList();
        if (sorted.Count == pageCount) return "";

        var parts = new List<string>();
        for (int i = 0; i < sorted.Count;)
        {
            int j = i;
            while (j + 1 < sorted.Count && sorted[j + 1] == sorted[j] + 1) j++;
            parts.Add(i == j ? $"{sorted[i] + 1}" : $"{sorted[i] + 1}-{sorted[j] + 1}");
            i = j + 1;
        }
        return string.Join(",", parts);
    }
}
