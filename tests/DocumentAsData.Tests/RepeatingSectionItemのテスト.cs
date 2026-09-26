using FluentAssertions;

namespace Marimo.DocumentAsData.Test;

public class RepeatingSectionItemのテスト
{
    [Fact]
    public void ContentControlsは別の明細にある同じTagを区別して読み取れます()
    {
        // Items配下の2つのw15:repeatingSectionItemは、どちらもItemNameを持ちます。
        // 同名でも明細の境界が異なるため、それぞれの商品名を取得します。
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Items
            .Select(it => it.ContentControls["ItemName"].Text)
            .Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void ContentControlsへの書き込みは別の明細にある同じTagを変更しません()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Items[0].ContentControls["ItemName"].Text = "変更後";

        document.RepeatingSections["Items"].Items
            .Select(it => it.ContentControls["ItemName"].Text)
            .Should().Equal("変更後", "商品B");
    }

    [Fact]
    public void ContentControlsは明細の外にある同じTagを含めません()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細の内外に同じTag.docx"));

        document.RepeatingSections["Items"].Items[0].ContentControls
            .Select(it => it.Text)
            .Should().Equal("商品A");
    }

    [Fact]
    public void ContentControlsは同じ明細内に同じTagが複数あると失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "1件の明細内に同じTag.docx"));

        var action = () => document.RepeatingSections["Items"].Items[0].ContentControls["ItemName"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void MergeFieldsは明細ごとに同じ名前を読み書きできます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Items[0].MergeFields["Code"].Text = "変更後";

        document.RepeatingSections["Items"].Items
            .Select(it => it.MergeFields["Code"].Text)
            .Should().Equal("変更後", "B001");
    }

    [Fact]
    public void CheckBoxesは明細ごとに同じTagを読み書きできます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Items[0].CheckBoxes["Agreement"].IsChecked = false;

        document.RepeatingSections["Items"].Items
            .Select(it => it.CheckBoxes["Agreement"].IsChecked)
            .Should().Equal(false, true);
    }

    [Fact(Skip = "明細内の値アクセスのGreen後、親セクションのXMLマッピングを無視しないことをレビューします。")]
    public void 明細内の文字列読み取りはXMLマッピングされた親セクションを拒否します()
    {
        // dataBindingは内側のItemNameではなく、外側のItemsのw:sdtPrにあります。
        // Wordは親の連結先から明細を作り直すため、内側だけを読む経路でも拒否します。
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Items[0].ContentControls["ItemName"].Text;

        action.Should().Throw<NotSupportedException>();
    }

    [Fact(Skip = "明細内の値アクセスのGreen後、親セクションのXMLマッピングを無視しないことをレビューします。")]
    public void 明細内の文字列書き込みはXMLマッピングされた親セクションを拒否します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Items[0].ContentControls["ItemName"].Text = "変更後";

        action.Should().Throw<NotSupportedException>();
    }

    [Fact(Skip = "明細内の文字列Content ControlのGreen後、日付選択への接続をレビューします。")]
    public void DatePickersは明細ごとに同じTagを読み書きできます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Items[0].DatePickers["DeliveryDate"].SelectedDateTime =
            new(2026, 10, 1, 0, 0, 0, TimeSpan.Zero);

        document.RepeatingSections["Items"].Items
            .Select(it => it.DatePickers["DeliveryDate"].SelectedDateTime)
            .Should().Equal(
                new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero),
                new DateTimeOffset(2026, 9, 27, 0, 0, 0, TimeSpan.Zero));
    }
}
