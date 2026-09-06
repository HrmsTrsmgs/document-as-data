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
}
