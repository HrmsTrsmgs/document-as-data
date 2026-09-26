using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前設定のテスト
{
    const string CustomerDataDocumentFilePath =
        @"TestData\コード生成\customerData.docx";
    const string SharedNameDocumentFilePath =
        @"TestData\コード生成\顧客.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";
    const string DatePickerDocumentFilePath =
        @"TestData\コード生成\日付選択ContentControl.docx";

    [Fact]
    public void NameMappingsは繰り返しセクションと明細データのプロパティ名へ適用されます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(
                @"TestData\コード生成\repeatingTemplate.docx",
                options =>
                {
                    options.NameMappings["Items"] = "Lines";
                    options.NameMappings["ItemName"] = "ProductName";
                }))
            .GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingTemplate.docx");
        dynamic documentAccessor = document;
        IEnumerable<dynamic> items = documentAccessor.Lines;

        items.Select(it => it.ProductName as object).Should().Equal("商品A");
    }

    [Fact]
    public void NameMappingsで変更した明細データのプロパティから元のTagへ書き込めます()
    {
        var assembly = GeneratedCodeInspection.AssemblyFrom(
            [
                .. GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\repeatingTemplate.docx",
                    options => options.NameMappings["ItemName"] = "ProductName"),
                """
                using System.Linq;
                namespace Generated;
                public static class UseMappedRepeatingItems
                {
                    public static string[] Replace(RepeatingTemplateDocument document)
                    {
                        document.Items.Replace([new() { ProductName = "商品C" }]);
                        return document.ContentControls.Select(it => it.Text).ToArray();
                    }
                }
                """
            ]);
        using var document = assembly.GeneratedType("RepeatingTemplateDocument")
            .InvokeStaticMethod<Document>("Open", @"TestData\コード生成\repeatingTemplate.docx");

        assembly.GeneratedType("UseMappedRepeatingItems").InvokeStaticMethod<string[]>("Replace", document)
            .Should().Equal("商品C");
    }

    [Fact]
    public void 生成時に同じ明細内で重複した名前は診断します()
    {
        GeneratedCodeInspection.GenerateDiagnostics(@"TestData\コード生成\repeatingDuplicateItem.docx")
            .Should().Contain(it => it.IsError && it.GeneratedName == "ItemName");
    }

    [Fact]
    public void 書式文字を含むTagから生成したDataプロパティへ文字列を読み込みます()
    {
        // TagはCustomer\u200CNameです。U+200Cは不可視の書式文字で、
        // C#ソースの識別子に含めてもコンパイル後の名前では無視されます。
        // 見た目や生成ソースだけでなく、Readから元のTagへ対応できることを確認します。
        const string filePath = @"TestData\コード生成\書式文字を含むTag.docx";
        using var document = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(filePath))
            .GeneratedType("書式文字を含むTagDocument")
            .InvokeStaticMethod<Document>("Open", filePath);
        dynamic documentAccessor = document;

        dynamic tested = documentAccessor.Read();

        (tested.CustomerName as object)
            .Should().BeOfType<string>()
            .Which.Should().Be("山田太郎");
    }

    [Fact]
    public void NameMappingsは自動名前変換より優先されます()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CustomerDataDocumentFilePath,
                    options =>
                        options.NameMappings["customerName"] = "ClientName"))
            .GeneratedType("CustomerDataDocument")
            .GetProperty("ClientName")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは辞書を代入して設定できます()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CustomerDataDocumentFilePath,
                    options =>
                        options.NameMappings = new()
                        {
                            ["customerName"] = "ClientName"
                        }))
            .GeneratedType("CustomerDataDocument")
            .GetProperty("ClientName")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは対象種類を指定せず同じ元名へ適用されます()
    {
        var assembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(
                SharedNameDocumentFilePath,
                options => options.NameMappings["顧客"] = "Customer"));

        var generatedType = assembly.GeneratedType("CustomerDocument");

        generatedType.GetProperty("Customer").Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは文字列ContentControlのTagへ適用されます()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath,
                    options =>
                        options.NameMappings["CustomerName"] = "ClientName"))
            .GeneratedType("文字列ContentControlDocument")
            .GetProperty("ClientName")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsはCheckBoxのTagへ適用されます()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath,
                    options =>
                        options.NameMappings["Agreement"] = "Consent"))
            .GeneratedType("チェック済みCheckBoxDocument")
            .GetProperty("Consent")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsはDatePickerのTagへ適用されます()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath,
                    options =>
                        options.NameMappings["DeliveryDate"] = "DueDate"))
            .GeneratedType("日付選択ContentControlDocument")
            .GetProperty("DueDate")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティへMERGEFIELDの文字列を読み込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CustomerDataDocumentFilePath,
                    options =>
                        options.NameMappings["customerName"] = "ClientName"))
            .GeneratedType("CustomerDataDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CustomerDataDocumentFilePath);
        dynamic documentAccessor = document;
        dynamic tested = documentAccessor.Read();

        (tested.ClientName as object)
            .Should().BeOfType<string>()
            .Which.Should().Be("株式会社○○");
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティからMERGEFIELDへ文字列を書き込みます()
    {
        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CustomerDataDocumentFilePath,
                    options =>
                        options.NameMappings["customerName"] = "ClientName"))
            .GeneratedType("CustomerDataDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CustomerDataDocumentFilePath);
        dynamic documentAccessor = tested;
        dynamic data = documentAccessor.Read();
        data.ClientName = "変更後";

        documentAccessor.Replace(data);

        tested.MergeFields["customerName"].Text.Should().Be("変更後");
    }

    [Fact]
    public void NameMappingsで変更した生成Dataプロパティへ文字列ContentControlの文字列を読み込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath,
                    options =>
                        options.NameMappings["CustomerName"] = "ClientName"))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);
        dynamic documentAccessor = document;
        dynamic tested = documentAccessor.Read();

        (tested.ClientName as object)
            .Should().BeOfType<string>()
            .Which.Should().Be("山田太郎");
    }

    [Fact]
    public void NameMappingsで変更した生成Dataプロパティから文字列ContentControlへ文字列を書き込みます()
    {
        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath,
                    options =>
                        options.NameMappings["CustomerName"] = "ClientName"))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);
        dynamic documentAccessor = tested;
        dynamic data = documentAccessor.Read();
        data.ClientName = "変更後";

        documentAccessor.Replace(data);

        tested.ContentControls["CustomerName"].Text.Should().Be("変更後");
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティへCheckBoxのチェック状態を読み込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath,
                    options =>
                        options.NameMappings["Agreement"] = "Consent"))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);
        dynamic documentAccessor = document;
        dynamic tested = documentAccessor.Read();

        (tested.Consent as object)
            .Should().BeOfType<bool>()
            .Which.Should().BeTrue();
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティからCheckBoxへチェック状態を書き込みます()
    {
        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath,
                    options =>
                        options.NameMappings["Agreement"] = "Consent"))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);
        dynamic documentAccessor = tested;
        dynamic data = documentAccessor.Read();
        data.Consent = false;

        documentAccessor.Replace(data);

        tested.CheckBoxes["Agreement"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティへDatePickerの選択日時を読み込みます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath,
                    options =>
                        options.NameMappings["DeliveryDate"] = "DueDate"))
            .GeneratedType("日付選択ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                DatePickerDocumentFilePath);
        dynamic documentAccessor = document;
        dynamic tested = documentAccessor.Read();

        (tested.DueDate as object)
            .Should().BeOfType<DateTimeOffset>()
            .Which.Should().Be(
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
    public void NameMappingsで変更した生成DataプロパティからDatePickerへ選択日時を書き込みます()
    {
        var selectedDateTime = new DateTimeOffset(
            2026,
            12,
            31,
            0,
            0,
            0,
            TimeSpan.FromHours(9));
        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath,
                    options =>
                        options.NameMappings["DeliveryDate"] = "DueDate"))
            .GeneratedType("日付選択ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                DatePickerDocumentFilePath);
        dynamic documentAccessor = tested;
        dynamic data = documentAccessor.Read();
        data.DueDate = selectedDateTime;

        documentAccessor.Replace(data);

        var written = tested.DatePickers["DeliveryDate"].SelectedDateTime;
        written.Should().Be(selectedDateTime);
        written.Offset.Should().Be(selectedDateTime.Offset);
    }
}
