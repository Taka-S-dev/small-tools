using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public class DialogWatchTests
{
    private static readonly DateTime T0 = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    private static readonly IntPtr Dialog = new(0x1234), Other = new(0x5678);

    private static DateTime At(double seconds) => T0.AddSeconds(seconds);
    private static TimeSpan Cpu(double ms) => TimeSpan.FromMilliseconds(ms);

    [Fact]
    public void Fires_once_when_a_dialog_stays_open_while_Office_does_nothing()
    {
        var judge = new DialogWatch.Judge();
        Assert.False(judge.Step(Dialog, Cpu(1000), At(0)));
        Assert.False(judge.Step(Dialog, Cpu(1010), At(4)));
        Assert.True(judge.Step(Dialog, Cpu(1020), At(5)));
        Assert.False(judge.Step(Dialog, Cpu(1020), At(9))); // 2 度は知らせない
    }

    [Fact]
    public void Does_not_fire_while_Office_is_working_behind_a_progress_dialog()
    {
        var judge = new DialogWatch.Judge();
        for (int second = 0; second <= 30; second++)
            Assert.False(judge.Step(Dialog, Cpu(second * 400), At(second)));
    }

    [Fact]
    public void Starts_counting_again_for_a_new_dialog_or_when_the_dialog_goes_away()
    {
        var judge = new DialogWatch.Judge();
        judge.Step(Dialog, Cpu(0), At(0));
        Assert.False(judge.Step(Other, Cpu(0), At(4)));     // 別の画面に替わった
        Assert.False(judge.Step(Other, Cpu(0), At(8)));     // 替わってから 4 秒
        Assert.False(judge.Step(IntPtr.Zero, Cpu(0), At(9)));
        Assert.False(judge.Step(Other, Cpu(0), At(10)));    // 消えてまた出た
        Assert.True(judge.Step(Other, Cpu(0), At(15)));
    }

    [Theory]
    [InlineData("#32770", true)]
    [InlineData("bosa_sdm_msword", true)]
    [InlineData("bosa_sdm_XL9", true)]
    [InlineData("NUIDialog", true)]
    [InlineData("OpusApp", false)]
    [InlineData("XLMAIN", false)]
    [InlineData("PPTFrameClass", false)]
    public void Recognises_dialog_window_classes(string className, bool expected) =>
        Assert.Equal(expected, DialogWatch.IsDialogClass(className));
}
