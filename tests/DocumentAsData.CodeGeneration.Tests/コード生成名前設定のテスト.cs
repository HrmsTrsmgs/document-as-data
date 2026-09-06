using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前設定のテスト
{
    const string CustomerDataDocumentFilePath =
        @"TestData\コード生成\customerData.docx";

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
}
