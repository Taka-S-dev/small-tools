using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public class OfficeKillTests
{
    [Theory]
    // 時間切れ・確認の画面で止まったとき（onlyIfHidden: false）
    [InlineData(true, false, false, false, true)]   // 隠れたまま止まった自分の Office は止める
    [InlineData(true, false, true, false, true)]    // 自分の Office が確認の画面だけ出して止まった
    [InlineData(true, false, true, true, false)]    // 利用者の文書が出ている。編集中の文書ごと消さない
    [InlineData(false, false, false, false, false)] // 利用者の Office を借りているだけなら止めない
    // アプリを閉じるとき（onlyIfHidden: true）
    [InlineData(true, true, false, false, true)]
    [InlineData(true, true, true, false, false)]
    [InlineData(true, true, true, true, false)]
    public void MayKill_never_stops_an_Office_that_shows_a_user_document(
        bool owned, bool onlyIfHidden, bool anyWindowVisible, bool userDocumentVisible, bool expected) =>
        Assert.Equal(expected, OfficeApp.MayKill(owned, onlyIfHidden, anyWindowVisible, userDocumentVisible));
}
