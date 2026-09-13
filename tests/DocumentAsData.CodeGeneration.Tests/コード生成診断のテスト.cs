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
    public void 生成プロパティ名が生成Document型のReadメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Read")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Read",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型のOpenメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Open")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Open",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型のReplaceメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Replace")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Replace",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのMergeFieldsコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "MergeFields")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "MergeFields",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのContentControlsコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\文字列ContentControl.docx",
                options => options.NameMappings["CustomerName"] = "ContentControls")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "ContentControls",
                    ["CustomerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのCheckBoxesコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\チェック済みCheckBox.docx",
                options => options.NameMappings["Agreement"] = "CheckBoxes")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CheckBoxes", ["Agreement"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのDatePickersコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\日付選択ContentControl.docx",
                options => options.NameMappings["DeliveryDate"] = "DatePickers")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "DatePickers", ["DeliveryDate"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型の項目検証メソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "ValidateRequiredItems")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "ValidateRequiredItems", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型名と同じ場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "CustomerDataDocument")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CustomerDataDocument", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Data型名と同じ場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "CustomerDataData")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CustomerDataData", ["customerName"]));
    }

    [Theory]
    [InlineData("OrderDocument")]
    [InlineData("OrderData")]
    public void NameMappingsで変更した生成型名とプロパティ名が同じ場合にもエラーを診断します(string propertyName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings = new()
                {
                    ["customerData"] = "Order",
                    ["customerName"] = propertyName
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, propertyName, ["customerName"]));
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
