using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string MergeFieldNameCollisionDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD名衝突.docx";
    const string DifferentItemTypeNameCollisionDocumentFilePath =
        @"TestData\コード生成\異種項目名衝突.docx";
    const string InvalidDocumentNameFilePath =
        @"TestData\コード生成\---.docx";
    const string CheckBoxAndDatePickerNameCollisionDocumentFilePath =
        @"TestData\コード生成\CheckBoxとDatePickerの名前衝突.docx";

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

    [Fact]
    public void NameMappingsで生成名を変更するとMERGEFIELDプロパティ名の衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options =>
                    options.NameMappings["customer-id"] = "CustomerIdDash")
            .Should().BeEmpty();
    }

    [Fact]
    public void 同じ生成Document型内の異なる種類のプロパティ名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(DifferentItemTypeNameCollisionDocumentFilePath)
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "CustomerId"
                && it.SourceNames.Contains("customer_id")
                && it.SourceNames.Contains("customer-id"));
    }

    [Fact]
    public void 区切り文字だけの文書名はCSharp識別子を生成できないため診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(InvalidDocumentNameFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "",
                    ["---"],
                    "---"));
    }

    [Fact]
    public void CheckBoxとDatePickerの生成プロパティ名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                CheckBoxAndDatePickerNameCollisionDocumentFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "DataItem",
                    ["data_item", "data-item"]));
    }
}
