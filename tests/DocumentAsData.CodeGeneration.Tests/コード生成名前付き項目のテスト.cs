using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前付き項目のテスト : IDisposable
{
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";
    const string BackslashMergeFieldNameDocumentFilePath =
        @"TestData\コード生成\バックスラッシュを含むMERGEFIELD名.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string SpecialCharacterTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTag.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";
    const string SpecialCharacterCheckBoxTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTagのCheckBox.docx";
    const string DatePickerDocumentFilePath =
        @"TestData\コード生成\日付選択ContentControl.docx";
    const string SpecialCharacterDatePickerTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTagのDatePicker.docx";

    readonly TemporaryDocumentFiles temporaryFiles = new();

    [Fact]
    public void 生成された繰り返し項目は明細が1件でもコレクションとして読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingTemplate.docx"))
            .GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingTemplate.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.ItemName as object).Should().Equal("商品A");
    }

    [Fact]
    public void 明細1件から生成した繰り返し項目で複数の明細を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingTemplate.docx"))
            .GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingItems.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.ItemName as object).Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void 複数の明細から生成しても同じ項目名を文書全体の重複として扱いません()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingItems.docx"))
            .GeneratedType("RepeatingItemsDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingItems.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.ItemName as object).Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void 生成された繰り返し項目へ明細データを書き込んで件数を変更できます()
    {
        var assembly = GeneratedCodeInspection.AssemblyFrom(
            [
                .. GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingTemplate.docx"),
                """
                using System.Linq;
                namespace Generated;
                public static class UseRepeatingItems
                {
                    public static string[] Replace(RepeatingTemplateDocument document)
                    {
                        document.Items.Replace([
                            new() { ItemName = "商品C" },
                            new() { ItemName = "商品D" }
                        ]);
                        return document.ContentControls.Select(it => it.Text).ToArray();
                    }
                }
                """
            ]);
        using var document = assembly.GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingTemplate.docx");

        assembly.GeneratedType("UseRepeatingItems").InvokeStaticMethod<string[]>("Replace", document)
            .Should().Equal("商品C", "商品D");
    }

    [Fact]
    public void 生成された通常項目と明細内の項目は同じ名前でも別々に読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingOutside.docx"))
            .GeneratedType("RepeatingOutsideDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingOutside.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        (documentAccessor.ItemName as object).Should().Be("文書全体の商品名");
        items.Select(it => it.ItemName as object).Should().Equal("商品A", "商品B");
    }

    [Fact]
    public void 生成された明細データの変更はReplaceするまで文書へ反映されません()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingTemplate.docx"))
            .GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingTemplate.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;
        var data = items.Single();

        data.ItemName = "変更後";

        document.ContentControls["ItemName"].Text.Should().Be("商品A");
    }

    [Fact]
    public void 生成された明細データでMERGEFIELDの文字列を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingKinds.docx"))
            .GeneratedType("RepeatingKindsDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingKinds.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.Code as object).Should().Equal("A001", "B001");
    }

    [Fact]
    public void 生成された明細データでチェック状態を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingKinds.docx"))
            .GeneratedType("RepeatingKindsDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingKinds.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.Agreement as object).Should().Equal(true, true);
    }

    [Fact(Skip = "型付き繰り返し項目の生成後、明細内日付選択のデータ生成への接続をレビューします。")]
    public void 生成された明細データで日時を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(@"TestData\コード生成\repeatingKinds.docx"))
            .GeneratedType("RepeatingKindsDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingKinds.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Items;

        items.Select(it => it.DeliveryDate as object).Should().Equal(
            new DateTimeOffset(2026, 9, 26, 0, 0, 0, TimeSpan.Zero),
            new DateTimeOffset(2026, 9, 27, 0, 0, 0, TimeSpan.Zero));
    }

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldsDocumentFilePath))
            .GeneratedType("MergefieldDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                MergeFieldsDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void MERGEFIELDから生成したプロパティは同名の文字列ContentControlから文字列を読み取れます()
    {
        // 生成元はCustomerNameだけを持つMERGEFIELD文書、開く文書は同名の文字列Content Controlです。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\単一項目.docx"))
            .GeneratedType("単一項目Document")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void MERGEFIELDから生成したプロパティは同名の文字列ContentControlよりMERGEFIELDを優先して読み取ります()
    {
        // 生成元はCustomerNameが1件、開く文書には同名のMERGEFIELDと文字列Content Controlが各1件あります。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\単一項目.docx"))
            .GeneratedType("単一項目Document")
            .InvokeStaticMethod<Document>(
                "Open",
                @"TestData\コード生成\同名のMERGEFIELDとContent Control.docx");

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("MERGEFIELDの値");
    }

    [Fact]
    public void バックスラッシュを含む名前から生成したMERGEFIELDプロパティで文字列を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BackslashMergeFieldNameDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp customer"] = "SpecialValue"))
            .GeneratedType("バックスラッシュを含むMERGEFIELD名Document")
            .InvokeStaticMethod<Document>(
                "Open",
                BackslashMergeFieldNameDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.SpecialValue;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void 改行を含む名前から生成したMERGEFIELDプロパティで文字列を読み取れます()
    {
        // 既存の固定DOCXをリンクして使用します。名前中のLFは表示用の改行ではなく、
        // w:instrText内で引用符に囲まれた名前Customer\nNameの一部です。
        // 生成コードをコンパイルし、元の名前のまま文書を開いて読み取れることを確認します。
        var filePath = @"TestData\コード生成\名前に改行を含むMERGEFIELD.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    filePath,
                    options => options.NameMappings["Customer\nName"] = "CustomerName"))
            .GeneratedType("名前に改行を含むMERGEFIELDDocument")
            .InvokeStaticMethod<Document>("Open", filePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void LFを含むTagから生成した文字列ContentControlプロパティで文字列を読み取れます()
    {
        // w:tagのval内の文字参照&#10;は、名前の一部であるLFです。
        // 表示内容の改行ではありません。生成コメントと文字列リテラルを壊さず、
        // 元のTagで項目へ到達できることを、生成コードのコンパイルと実行で確認します。
        var filePath = @"TestData\コード生成\LFを含むTagの文字列ContentControl.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    filePath,
                    options => options.NameMappings["Customer\nName"] = "CustomerName"))
            .GeneratedType("LFを含むTagの文字列ContentControlDocument")
            .InvokeStaticMethod<Document>("Open", filePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void LFを含むTagから生成したCheckBoxプロパティでチェック状態を読み取れます()
    {
        // w:tagのvalにLFを含みます。CheckBoxには文字列Content Controlとは別の
        // Document/Dataプロパティの生成処理があるため、その経路を確認します。
        var filePath = @"TestData\コード生成\LFを含むTagのCheckBox.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    filePath,
                    options => options.NameMappings["Customer\nName"] = "Agreement"))
            .GeneratedType("LFを含むTagのCheckBoxDocument")
            .InvokeStaticMethod<Document>("Open", filePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.Agreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void LFを含むTagから生成したDatePickerプロパティで日時を読み取れます()
    {
        // w:tagのvalにLFを含みます。日付の表示形式ではなく、
        // DatePickerのDocument/DataプロパティへTagを埋め込む経路を確認します。
        var filePath = @"TestData\コード生成\LFを含むTagのDatePicker.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    filePath,
                    options => options.NameMappings["Customer\nName"] = "DeliveryDate"))
            .GeneratedType("LFを含むTagのDatePickerDocument")
            .InvokeStaticMethod<Document>("Open", filePath);

        dynamic documentAccessor = document;
        DateTimeOffset tested = documentAccessor.DeliveryDate;

        tested.Should().Be(new DateTimeOffset(2026, 9, 4, 0, 0, 0, TimeSpan.Zero));
    }

    [Fact]
    public void CRを含むTagから生成した文字列ContentControlプロパティで文字列を読み取れます()
    {
        // w:tagのvalの文字参照&#13;はCRです。LFとは異なり、
        // 現状は生成コメントだけでなくC#の文字列リテラルにも生の改行が入ります。
        var filePath = @"TestData\コード生成\CRを含むTagの文字列ContentControl.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    filePath,
                    options => options.NameMappings["Customer\rName"] = "CustomerName"))
            .GeneratedType("CRを含むTagの文字列ContentControlDocument")
            .InvokeStaticMethod<Document>("Open", filePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           MergeFieldsDocumentFilePath))
                   .GeneratedType("MergefieldDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       MergeFieldsDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.MergeFields["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void MERGEFIELDから生成したプロパティは同名の文字列ContentControlへ文字列を書き込めます()
    {
        // 生成元はMERGEFIELDですが、開く文書では同名の文字列Content Controlです。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\単一項目.docx"))
            .GeneratedType("単一項目Document")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);

        dynamic documentAccessor = document;
        documentAccessor.CustomerName = "生成後";

        document.ContentControls["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void MERGEFIELDから生成したプロパティは同名の文字列ContentControlを変更せずMERGEFIELDへ書き込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\単一項目.docx"))
            .GeneratedType("単一項目Document")
            .InvokeStaticMethod<Document>(
                "Open",
                @"TestData\コード生成\同名のMERGEFIELDとContent Control.docx");

        // CustomerNameの両種類が各1件ありますが、生成元と同じMERGEFIELDだけを書き換えます。
        dynamic documentAccessor = document;
        documentAccessor.CustomerName = "変更後";

        document.MergeFields["CustomerName"].Text.Should().Be("変更後");
        document.ContentControls["CustomerName"].Text.Should().Be("Content Controlの値");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 文字列ContentControlから生成したプロパティは同名のMERGEFIELDから文字列を読み取れます()
    {
        // 生成元はTagがCustomerNameの文字列Content Controlですが、開く文書では同名のMERGEFIELDです。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                MergeFieldsDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void 文字列ContentControlから生成したプロパティは同名のMERGEFIELDより文字列ContentControlを優先して読み取ります()
    {
        // 生成元はCustomerNameが1件、開く文書には同名のMERGEFIELDと文字列Content Controlが各1件あります。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                @"TestData\コード生成\同名のMERGEFIELDとContent Control.docx");

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("Content Controlの値");
    }

    [Fact]
    public void 特殊文字を含むTagから生成したプロパティで文字列を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""document"] = "SpecialValue"))
            .GeneratedType("引用符とバックスラッシュを含むTagDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterTagDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.SpecialValue;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           TextContentControlDocumentFilePath))
                   .GeneratedType("文字列ContentControlDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       TextContentControlDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.ContentControls["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 文字列ContentControlから生成したプロパティは同名のMERGEFIELDへ文字列を書き込めます()
    {
        // 生成元は文字列Content Controlですが、開く文書では同名のMERGEFIELDです。
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                MergeFieldsDocumentFilePath);

        dynamic documentAccessor = document;
        documentAccessor.CustomerName = "生成後";

        document.MergeFields["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 文字列ContentControlから生成したプロパティは同名のMERGEFIELDを変更せず文字列ContentControlへ書き込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                @"TestData\コード生成\同名のMERGEFIELDとContent Control.docx");

        // CustomerNameの両種類が各1件ありますが、生成元と同じ文字列Content Controlだけを書き換えます。
        dynamic documentAccessor = document;
        documentAccessor.CustomerName = "変更後";

        document.ContentControls["CustomerName"].Text.Should().Be("変更後");
        document.MergeFields["CustomerName"].Text.Should().Be("MERGEFIELDの値");
    }

    [Fact]
    public void 生成されたDocument型のCheckBoxプロパティからチェック状態を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.Agreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 生成されたDocument型のReadでCheckBoxを文書データとして読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);

        dynamic documentAccessor = document;
        var data = documentAccessor.Read();
        bool tested = data.Agreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 特殊文字を含むTagから生成したCheckBoxプロパティでチェック状態を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterCheckBoxTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""agreement"] = "SpecialAgreement"))
            .GeneratedType("引用符とバックスラッシュを含むTagのCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterCheckBoxTagDocumentFilePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.SpecialAgreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 生成されたDocument型のCheckBoxプロパティへチェック状態を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           CheckBoxDocumentFilePath))
                   .GeneratedType("チェック済みCheckBoxDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       CheckBoxDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.Agreement = false;
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.CheckBoxes["Agreement"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void 生成されたDocument型のReplaceで文書データをCheckBoxへ書き込めます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);
        dynamic documentAccessor = document;
        var data = documentAccessor.Read();

        data.Agreement = false;
        documentAccessor.Replace(data);

        document.CheckBoxes["Agreement"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void 生成されたDocument型のDatePickerプロパティから日時を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath))
            .GeneratedType("日付選択ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                DatePickerDocumentFilePath);

        dynamic documentAccessor = document;
        DateTimeOffset tested = documentAccessor.DeliveryDate;

        tested.Should().Be(
            new DateTimeOffset(
                2026,
                9,
                4,
                0,
                0,
                0,
                TimeSpan.Zero));
    }

    [Fact]
    public void 特殊文字を含むTagから生成したDatePickerプロパティで日時を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterDatePickerTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""delivery"] = "SpecialDeliveryDate"))
            .GeneratedType("引用符とバックスラッシュを含むTagのDatePickerDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterDatePickerTagDocumentFilePath);

        dynamic documentAccessor = document;
        DateTimeOffset tested = documentAccessor.SpecialDeliveryDate;

        tested.Should().Be(
            new DateTimeOffset(
                2026,
                9,
                4,
                0,
                0,
                0,
                TimeSpan.Zero));
    }

    [Fact]
    public void 生成されたDocument型のDatePickerプロパティへ日時を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();
        var value = new DateTimeOffset(
            2026,
            12,
            31,
            0,
            0,
            0,
            TimeSpan.FromHours(9));

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DatePickerDocumentFilePath))
                   .GeneratedType("日付選択ContentControlDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       DatePickerDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.DeliveryDate = value;
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);
        var actual = tested.DatePickers["DeliveryDate"].SelectedDateTime;

        actual.Should().Be(value);
        actual.Offset.Should().Be(value.Offset);
    }
}
