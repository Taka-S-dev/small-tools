using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public sealed class SlideTitlesTests
{
    [Theory]
    [InlineData("スライド 3: まとめ", "まとめ")]
    [InlineData("スライド12：売上の推移", "売上の推移")]
    [InlineData("Slide 1: Overview", "Overview")]
    [InlineData("スライド 4", "スライド 4")]
    [InlineData("スライドの作り方", "スライドの作り方")]
    [InlineData("第3章: まとめ", "第3章: まとめ")]
    public void Strip_drops_the_slide_number(string title, string expected) =>
        Assert.Equal(expected, SlideTitles.Strip(title));

    [Fact]
    public void Strip_applies_to_every_level()
    {
        OutlineNode[] outline = [new("スライド 1: 表紙", 0, [new("スライド 2: 目的", 1, [])])];
        var stripped = SlideTitles.Strip(outline);
        Assert.Equal("表紙", stripped[0].Title);
        Assert.Equal("目的", stripped[0].Children[0].Title);
    }
}
