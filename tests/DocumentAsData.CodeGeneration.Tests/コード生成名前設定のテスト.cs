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
}
