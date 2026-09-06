using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前付き項目のテスト
{
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldsDocumentFilePath))
            .GeneratedType("MERGEFIELDDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                MergeFieldsDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("株式会社○○");
    }
}
