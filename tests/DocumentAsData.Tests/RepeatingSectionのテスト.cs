using System.Runtime.CompilerServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Marimo.DocumentAsData.Test.テスト補助;
using Word2013 = DocumentFormat.OpenXml.Office2013.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class RepeatingSectionのテスト
{
    [Fact]
    public void 削除した明細のMERGEFIELDは参照を手放すと回収されます()
    {
        using var document = Document.Open(Path.Combine("TestData", "明細に4種類の項目.docx"));
        var field = 削除した明細のMERGEFIELDを取得する(document.RepeatingSections["Items"]);

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        field.IsAlive.Should().BeFalse();
        GC.KeepAlive(document);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    static WeakReference 削除した明細のMERGEFIELDを取得する(RepeatingSection section)
    {
        section.Replace<CodeData>([new() { Code = "A" }, new() { Code = "B" }]);
        var field = new WeakReference(section.Items[1].MergeFields["Code"]);
        section.Replace<CodeData>([new() { Code = "A" }]);
        return field;
    }

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

    [Fact]
    public void Replaceは属性で指定した明細内のTagへ書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<NamedItemData>(
            [new() { Name = "商品C" }, new() { Name = "商品D" }]);

        document.ContentControls.Select(it => it.Text).Should().Equal("商品C", "商品D");
    }

    [Fact]
    public void Replaceは明細内のMERGEFIELDへ対応するプロパティを書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Replace<CodeData>(
            [new() { Code = "C001" }, new() { Code = "D001" }]);

        document.RepeatingSections["Items"].Items.Select(it => it.MergeFields["Code"].Text)
            .Should().Equal("C001", "D001");
    }

    [Fact]
    public void Replaceは明細内のチェックボックスへboolプロパティを書き込みます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "明細に4種類の項目.docx"));

        document.RepeatingSections["Items"].Replace<AgreementData>(
            [new() { Agreement = false }, new() { Agreement = true }]);

        document.RepeatingSections["Items"].Items.Select(it => it.CheckBoxes["Agreement"].IsChecked)
            .Should().Equal(false, true);
    }

    [Fact]
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

    [Fact]
    public void Replaceは入力データに合わせて明細を増やします()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);

        document.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName).Should().Equal("商品C", "商品D");
    }

    [Fact]
    public void Replaceは入力データに合わせて明細を減らします()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        document.RepeatingSections["Items"].Replace<ItemData>([new() { ItemName = "商品C" }]);

        document.RepeatingSections["Items"].Read<ItemData>()
            .Select(it => it.ItemName).Should().Equal("商品C");
    }

    [Fact]
    public void Replace後も同じセクションから明細を取得して再度変更できます()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));
        var tested = document.RepeatingSections["Items"];

        tested.Replace<ItemData>([new() { ItemName = "商品C" }, new() { ItemName = "商品D" }]);
        tested.Items[1].ContentControls["ItemName"].Text = "商品E";

        tested.Read<ItemData>().Select(it => it.ItemName).Should().Equal("商品C", "商品E");
    }

    [Fact]
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

    [Fact]
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
        new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(saved, TestCancellation.Token).Should().BeEmpty();
    }

    [Fact]
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

    [Fact]
    public void Replaceで複数の明細を増やすと識別IDは整数の上限から折り返します()
    {
        var inputPath = Path.Combine("TestData", "識別IDが整数の上下限を含む2件の明細.docx");

        // 固定データは「表行を繰り返す2件の明細.docx」のw:idを2か所だけ上下限へ変更したものです。
        // Wordによる生成差異ではなく、既存IDの境界値を含む文書での複製を検証します。
        ReadValidatedContentControlIds(inputPath)
            .Should().Contain(int.MaxValue).And.Contain(int.MinValue).And.OnlyHaveUniqueItems();

        using var output = new MemoryStream();
        using (var document = Document.Open(inputPath))
        {
            document.RepeatingSections["Items"].Items.Should().HaveCount(2);
            document.RepeatingSections["Items"].Replace<ItemData>(
            [
                new() { ItemName = "商品C" },
                new() { ItemName = "商品D" },
                new() { ItemName = "商品E" },
                new() { ItemName = "商品F" }
            ]);
            document.SaveAs(output);
        }

        ReadSavedItemNames(output).Should().Equal("商品C", "商品D", "商品E", "商品F");

        // 複製のたびに元の明細の直後へ挿入するので、後から採番した複製が文書順では先になります。
        ReadValidatedContentControlIds(output).Should().Equal(
            int.MaxValue, int.MinValue, 143, 146, 144,
            int.MinValue + 3, int.MinValue + 4, int.MinValue + 1, int.MinValue + 2);
    }

    [Fact]
    public void ReadはXMLマッピングされたセクションを拒否します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Read<ItemData>().ToArray();

        action.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void ReplaceはXMLマッピングされたセクションを拒否します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "XMLマッピングされた繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"].Replace<ItemData>(
            [new() { ItemName = "商品C" }]);

        action.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void Replaceは0件の入力を拒否して既存の明細を保持します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));
        var tested = document.RepeatingSections["Items"];

        var action = () => tested.Replace<ItemData>([]);

        action.Should().ThrowExactly<ArgumentException>().Which.ParamName.Should().Be("items");
        tested.Read<ItemData>().Select(it => it.ItemName).Should().Equal("商品A", "商品B");
    }

    /// <summary>
    /// 保存先を先頭から読み直し、文書を閉じる前に明細の商品名を取り出します。
    /// </summary>
    static string[] ReadSavedItemNames(Stream output)
    {
        output.Position = 0;
        using var saved = Document.Open(output);
        return [.. saved.RepeatingSections["Items"].Read<ItemData>().Select(it => it.ItemName)];
    }

    /// <summary>
    /// 固定データを開き、Open XMLとしての妥当性とContent Controlの識別IDを確認します。
    /// </summary>
    static int?[] ReadValidatedContentControlIds(string filePath)
    {
        using var input = WordprocessingDocument.Open(filePath, false);
        return ReadValidatedContentControlIds(input);
    }

    /// <summary>
    /// 保存先を先頭から開き直し、Open XMLとしての妥当性とContent Controlの識別IDを確認します。
    /// </summary>
    static int?[] ReadValidatedContentControlIds(Stream output)
    {
        output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        return ReadValidatedContentControlIds(saved);
    }

    /// <summary>
    /// w:idは公開APIのTagとは別の識別子です。値の読み戻しやValidatorだけでは
    /// 重複を観測できないため、文書を閉じる前にOOXMLから取り出します。
    /// </summary>
    static int?[] ReadValidatedContentControlIds(WordprocessingDocument package)
    {
        new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package).Should().BeEmpty();
        var document = package.MainDocumentPart?.Document;
        document.Should().NotBeNull();
        return [.. document.Descendants<Wordprocessing.SdtId>().Select(it => it.Val?.Value)];
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
