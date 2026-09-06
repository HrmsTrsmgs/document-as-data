using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string MergeFieldNameCollisionDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD名衝突.docx";

    [Fact]
    public void 自動変換後に同じMERGEFIELDプロパティ名となる場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(MergeFieldNameCollisionDocumentFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "CustomerId",
                    ["customer_id", "customer-id"]));
    }

    [Fact]
    public void MERGEFIELDプロパティ名の衝突を自動的な連番追加では解消しません()
    {
        GeneratedCodeInspection
            .GenerateSources(MergeFieldNameCollisionDocumentFilePath)
            .Should().NotContain(
                source => source.Contains(
                    "CustomerId2",
                    StringComparison.Ordinal));
    }
}
