using DocsToPdf.Core;

namespace DocsToPdf.Tests;

public class SettingChangesTests
{
    /// <summary>Office の設定の代わり。名前ごとの値を持ち、書けない設定を作れる。</summary>
    private sealed class FakeSettings
    {
        public Dictionary<string, int> Values { get; } = new() { ["alerts"] = -1, ["security"] = 1, ["events"] = 1 };
        public HashSet<string> ReadOnly { get; } = [];
        public List<string> Writes { get; } = [];

        public void Change(OfficeApp.SettingChanges c, string name, int value) =>
            c.Change(() => Values[name], v =>
            {
                if (ReadOnly.Contains(name)) throw new InvalidOperationException($"{name} は変えられない");
                Writes.Add($"{name}={v}");
                Values[name] = v;
            }, value);
    }

    [Fact]
    public void Restores_every_setting_in_reverse_order()
    {
        var office = new FakeSettings();
        Action restore = OfficeApp.SettingChanges.Apply(c =>
        {
            office.Change(c, "alerts", 0);
            office.Change(c, "security", 3);
        });
        Assert.Equal(3, office.Values["security"]);

        restore();

        Assert.Equal(-1, office.Values["alerts"]);
        Assert.Equal(1, office.Values["security"]);
        Assert.Equal(["alerts=0", "security=3", "security=1", "alerts=-1"], office.Writes);
    }

    [Fact]
    public void Undoes_the_settings_already_changed_when_a_later_one_fails()
    {
        var office = new FakeSettings();
        office.ReadOnly.Add("security");

        Assert.Throws<InvalidOperationException>(() => OfficeApp.SettingChanges.Apply(c =>
        {
            office.Change(c, "alerts", 0);
            office.Change(c, "security", 3);
        }));

        // マクロを無効にできなかったので、確認を止めた設定も元に戻っている
        Assert.Equal(-1, office.Values["alerts"]);
        Assert.Equal(1, office.Values["security"]);
    }

    [Fact]
    public void Keeps_restoring_after_one_setting_cannot_be_written_back()
    {
        var office = new FakeSettings();
        Action restore = OfficeApp.SettingChanges.Apply(c =>
        {
            office.Change(c, "events", 0);
            office.Change(c, "alerts", 0);
        });
        office.ReadOnly.Add("alerts");

        restore();

        Assert.Equal(1, office.Values["events"]);
    }
}
