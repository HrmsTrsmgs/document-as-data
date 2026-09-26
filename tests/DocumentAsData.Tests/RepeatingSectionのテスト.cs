using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Word2013 = DocumentFormat.OpenXml.Office2013.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class RepeatingSectionのテスト
{
    [Fact]
    public void Itemsは明細が1件でもコレクションとして取得できます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        document.RepeatingSections["Items"].Items.Should().ContainSingle();
    }

    [Fact]
    public void Readは明細ごとに自作クラスへ対応付けて文書順に読み取ります()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細の内外に同じTag.docx"));

        document.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName)
            .Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void Readは属性で指定した明細内のTagを使用します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Read<NamedItemData>()
            .Select(it => it.Name)
            .Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void Readはある明細で必要な項目が不足すると失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションの明細に項目が不足.docx"));

        var action = () => document.RepeatingSections["Items"].Read<ItemData>().ToArray();

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void Replaceは同じ件数の明細を順番に変更しセクション外の項目を保持します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細の内外に同じTag.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);

        document.ContentControls.Select(it => it.Text)
            .Should().Equal("文書全体の商品名", "商品C", "商品D");
    }

    [Fact(Skip = "同件数のReplaceのGreen後、文書項目名を指定する既存属性との接続をレビューします。")]
    public void Replaceは属性で指定した明細内のTagへ書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<NamedItemData>(
            [new() { Name = "商品C" }, new() { Name = "商品D" }]);

        document.ContentControls.Select(it => it.Text).Should().Equal("商品C", "商品D");
    }

    [Fact(Skip = "明細内MERGEFIELDのGreen後、自作クラスからの書き込みへの接続をレビューします。")]
    public void Replaceは明細内のMERGEFIELDへ対応するプロパティを書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Replace<CodeData>(
            [new() { Code = "C001" }, new() { Code = "D001" }]);

        document.RepeatingSections["Items"].Items.Select(it => it.MergeFields["Code"].Text)
            .Should().Equal("C001", "D001");
    }

    [Fact(Skip = "明細内チェックボックスのGreen後、自作クラスからの書き込みへの接続をレビューします。")]
    public void Replaceは明細内のチェックボックスへboolプロパティを書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Replace<AgreementData>(
            [new() { Agreement = false }, new() { Agreement = true }]);

        document.RepeatingSections["Items"].Items.Select(it => it.CheckBoxes["Agreement"].IsChecked)
            .Should().Equal(false, true);
    }

    [Fact(Skip = "明細内日付選択のGreen後、自作クラスからの書き込みへの接続をレビューします。")]
    public void Replaceは明細内の日付選択へ日時プロパティを書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Replace<DeliveryData>(
            [
                new() { DeliveryDate = new(2026, 10, 1, 0, 0, 0, TimeSpan.Zero) },
                new() { DeliveryDate = new(2026, 10, 2, 0, 0, 0, TimeSpan.Zero) }
            ]);

        document.RepeatingSections["Items"].Items.Select(it => it.DatePickers["DeliveryDate"].SelectedDateTime)
            .Should().Equal(
                new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero),
                new DateTimeOffset(2026, 10, 2, 0, 0, 0, TimeSpan.Zero));
    }

    [Fact(Skip = "同件数のReplaceのGreen後、既存の明細をもとに件数を増やす仕様をレビューします。")]
    public void Replaceは入力データに合わせて明細を増やします()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);

        document.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName).Should().Equal("商品C", "商品D");
    }

    [Fact(Skip = "同件数のReplaceのGreen後、不要な明細を除いて件数を減らす仕様をレビューします。")]
    public void Replaceは入力データに合わせて明細を減らします()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>([new() { ItemName = "商品C" }]);

        document.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName).Should().Equal("商品C");
    }

    [Fact(Skip = "件数変更のGreen後、取得済みセクションから変更後の明細を扱えることをレビューします。")]
    public void Replace後も同じセクションから明細を取得して再度変更できます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));
        var tested = document.RepeatingSections["Items"];

        tested.Replace<ItemData>([new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);
        tested.Items[1].ContentControls["ItemName"].Text = "商品E";

        tested.Read<ItemData>().Select(it => it.ItemName).Should().Equal("商品C", "商品E");
    }

    [Fact(Skip = "件数変更のGreen後、保存して開き直した文書でも明細を扱えることをレビューします。")]
    public void Replaceした明細の件数と値は保存後も保持されます()
    {
        using var output = new MemoryStream();
        using (var document = Document.Open(
                   Path.Combine("TestData", "繰り返しセクションのContent Control.docx")))
        {
            document.RepeatingSections["Items"].Replace<ItemData>(
                [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);
            document.SaveAs(output);
        }

        output.Position = 0;
        using var saved = Document.Open(output);

        saved.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName).Should().Equal("商品C", "商品D");
    }

    [Fact(Skip = "件数変更のGreen後、Wordで編集できる表行構造と書式の保持をレビューします。")]
    public void Replaceで増やした明細は表行の構造と書式を保持します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "表行を繰り返す2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }, new() { ItemName = "商品E" }]);

        using var output = new MemoryStream();
        document.SaveAs(output);
        output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        var body = saved.MainDocumentPart?.Document?.Body;

        body.Should().NotBeNull();

        // w:trは表の行。見出し1行と明細3行が残り、各明細のw:rPr/w:b（太字）も保持します。
        // w15:repeatingSectionItemは明細1件の枠で、行だけを複製してこの枠を失ってもいけません。
        // 表の後ろにある段落は繰り返しの対象外です。
        body.Descendants<Wordprocessing.TableRow>().Should().HaveCount(4);
        body.Descendants<Word2013.SdtRepeatedSectionItem>().Should().HaveCount(3);
        body.Descendants<Wordprocessing.Bold>().Should().HaveCount(3);
        body.Descendants<Wordprocessing.Text>().Select(it => it.Text)
            .Should().Equal("見出し", "商品C", "商品D", "商品E", "後書き");
        new OpenXmlValidator(FileFormatVersions.Office2013).Validate(saved).Should().BeEmpty();
    }

    [Fact(Skip = "件数変更のGreen後、複製されたContent ControlをWordが区別できることをレビューします。")]
    public void Replaceで増やした明細のContentControlは識別IDが重複しません()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "表行を繰り返す2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }, new() { ItemName = "商品E" }]);

        using var output = new MemoryStream();
        document.SaveAs(output);
        output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);

        // w:idはWordがContent Control自体を識別するIDです。
        // 各明細で同じTagを使うことは許しますが、複製元のw:idまで使い回してはいけません。
        var savedDocument = saved.MainDocumentPart?.Document;
        savedDocument.Should().NotBeNull();

        savedDocument.Descendants<Wordprocessing.SdtId>()
            .Select(it => it.Val?.Value)
            .Should().OnlyHaveUniqueItems();
    }

    [Fact(Skip = "明細のReadのGreen後、XMLマッピング済みセクションの拒否をレビューします。")]
    public void ReadはXMLマッピングされたセクションを拒否します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Read<ItemData>().ToArray();

        action.Should().Throw<NotSupportedException>();
    }

    [Fact(Skip = "明細のReplaceのGreen後、XMLマッピング済みセクションの拒否をレビューします。")]
    public void ReplaceはXMLマッピングされたセクションを拒否します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }]);

        action.Should().Throw<NotSupportedException>();
    }

    [Fact(Skip = "仕様未決定：0件にしたときの文書表現、再追加の複製元、保存後の再利用方法を決めてからテスト本体を書きます。")]
    public void Replaceへ0件を渡す場合の契約を決める必要があります()
    {
        // 空の明細を残すのか、別にテンプレートを保持するのか、0件を拒否するのか未決定です。
        // 期待結果を捏造しないため、この1件は仕様テストではなく検討メモとして数えます。
        throw new NotImplementedException();
    }

    /// <summary>
    /// 明細内のTagと同名のプロパティで読み書きする利用側のデータ型です。
    /// </summary>
    public sealed class ItemData
    {
        /// <summary>
        /// 明細の商品名を取得または設定します。
        /// </summary>
        public string ItemName { get; set; } = "";
    }

    /// <summary>
    /// 明細内のTagと異なるプロパティ名を使う利用側のデータ型です。
    /// </summary>
    public sealed class NamedItemData
    {
        /// <summary>
        /// ItemNameに対応する商品名を取得または設定します。
        /// </summary>
        [DocumentItemName("ItemName")]
        public string Name { get; set; } = "";
    }

    /// <summary>
    /// 明細内のMERGEFIELDへ対応付ける利用側のデータ型です。
    /// </summary>
    public sealed class CodeData
    {
        /// <summary>
        /// 明細の商品コードを取得または設定します。
        /// </summary>
        public string Code { get; set; } = "";
    }

    /// <summary>
    /// 明細内のチェックボックスへ対応付ける利用側のデータ型です。
    /// </summary>
    public sealed class AgreementData
    {
        /// <summary>
        /// 明細の確認状態を取得または設定します。
        /// </summary>
        public bool Agreement { get; set; }
    }

    /// <summary>
    /// 明細内の日付選択へ対応付ける利用側のデータ型です。
    /// </summary>
    public sealed class DeliveryData
    {
        /// <summary>
        /// 明細の納期日時を取得または設定します。
        /// </summary>
        public DateTimeOffset DeliveryDate { get; set; }
    }
}
