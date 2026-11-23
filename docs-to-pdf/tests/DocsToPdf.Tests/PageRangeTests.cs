using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public class PageRangeTests
{
    [Theory]
    [InlineData("１－３、５", "1-3,5")]
    [InlineData(" 1 - 3 , 5 ", "1-3,5")]
    [InlineData("8〜", "8-")]
    [InlineData(",2,", "2")]
    [InlineData("1-3 10-12", "1-3,10-12")]
    [InlineData("１－３　１０－１２", "1-3,10-12")]
    [InlineData("2 5 8-", "2,5,8-")]
    [InlineData(null, "")]
    [InlineData("   ", "")]
    public void Normalize_converts_full_width_and_removes_spaces(string? input, string expected) =>
        Assert.Equal(expected, PageRange.Normalize(input));

    [Theory]
    [InlineData("", true)]
    [InlineData("1-3,5,8-", true)]
    [InlineData("-3", true)]
    [InlineData("0", false)]
    [InlineData("3-1", false)]
    [InlineData("1-2-3", false)]
    [InlineData("-", false)]
    [InlineData("a", false)]
    public void IsValid(string input, bool expected) => Assert.Equal(expected, PageRange.IsValid(input));

    [Theory]
    [InlineData("1-3,5", 10, new[] { 0, 1, 2, 4 })]
    [InlineData("8-", 10, new[] { 7, 8, 9 })]
    [InlineData("-2", 10, new[] { 0, 1 })]
    [InlineData("5,1,5", 10, new[] { 0, 4 })]
    [InlineData("2,20", 3, new[] { 1 })]
    public void Resolve_returns_sorted_zero_based_pages(string input, int pageCount, int[] expected) =>
        Assert.Equal(expected, PageRange.Resolve(input, pageCount));

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void Resolve_uses_every_page_when_empty(string input) =>
        Assert.Equal([0, 1, 2], PageRange.Resolve(input, 3));

    [Theory]
    [InlineData("3-1")]
    [InlineData("abc")]
    [InlineData("1--3")]
    public void Resolve_returns_no_page_for_an_unreadable_range(string input) =>
        Assert.Empty(PageRange.Resolve(input, 3)); // 確かめ忘れても、全ページを出してしまわない

    [Fact]
    public void Resolve_returns_nothing_when_all_pages_are_out_of_range() =>
        Assert.Empty(PageRange.Resolve("5-", 3));

    [Theory]
    [InlineData(new[] { 0, 1, 2, 4, 7, 8 }, 10, "1-3,5,8-9")]
    [InlineData(new[] { 2, 0, 1 }, 3, "")]
    [InlineData(new[] { 3 }, 10, "4")]
    public void Format_writes_the_shortest_ranges(int[] pages, int pageCount, string expected) =>
        Assert.Equal(expected, PageRange.Format(pages, pageCount));

    [Theory]
    [InlineData("1-3,5,8-9", 10)]
    [InlineData("2,4,6", 6)]
    public void Format_and_Resolve_round_trip(string text, int pageCount) =>
        Assert.Equal(text, PageRange.Format(PageRange.Resolve(text, pageCount), pageCount));
}
