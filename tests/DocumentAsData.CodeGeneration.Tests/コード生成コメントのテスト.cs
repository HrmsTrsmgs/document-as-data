using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成コメントのテスト
{
    const string BasicStructureDocumentFilePath =
        @"TestData\コード生成\BasicStructure.docx";
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";
    const string DatePickerDocumentFilePath =
        @"TestData\コード生成\日付選択ContentControl.docx";

    [Fact]
    public void Document型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .SummaryText()
            .Should().Be("Word文書「BasicStructure」を型付きで表します。");
    }

    [Fact]
    public void Data型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureData")
            .SummaryText()
            .Should().Be("Word文書「BasicStructure」のデータを表します。");
    }

    [Fact]
    public void ファイルパスから開くOpenメソッドのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .MethodDeclaration("Open", "string")
            .SummaryText()
            .Should().Be("指定したファイルパスのWord文書を型付きで開きます。");
    }

    [Fact]
    public void Streamから開くOpenメソッドのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .MethodDeclaration("Open", "System.IO.Stream")
            .SummaryText()
            .Should().Be("指定したStream上のWord文書を型付きで開きます。");
    }

    [Fact]
    public void Readメソッドのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .MethodDeclaration("Read")
            .SummaryText()
            .Should().Be("Word文書全体のデータを読み込みます。");
    }

    [Fact]
    public void Replaceメソッドのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .MethodDeclaration("Replace", "BasicStructureData")
            .SummaryText()
            .Should().Be("Word文書全体のデータを置換します。");
    }

    [Fact]
    public void MERGEFIELDプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(MergeFieldsDocumentFilePath)
            .TypeDeclaration("MergefieldDocument")
            .PropertyDeclaration("CustomerName")
            .SummaryText()
            .Should().Be(
                "MERGEFIELD「CustomerName」の文字列を取得または設定します。");
    }

    [Fact]
    public void 文字列ContentControlプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(TextContentControlDocumentFilePath)
            .TypeDeclaration("文字列ContentControlDocument")
            .PropertyDeclaration("CustomerName")
            .SummaryText()
            .Should().Be(
                "文字列Content Control「CustomerName」の文字列を取得または設定します。");
    }

    [Fact]
    public void CheckBoxプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(CheckBoxDocumentFilePath)
            .TypeDeclaration("チェック済みCheckBoxDocument")
            .PropertyDeclaration("Agreement")
            .SummaryText()
            .Should().Be(
                "CheckBox「Agreement」のチェック状態を取得または設定します。");
    }

    [Fact]
    public void DatePickerプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DatePickerDocumentFilePath)
            .TypeDeclaration("日付選択ContentControlDocument")
            .PropertyDeclaration("DeliveryDate")
            .SummaryText()
            .Should().Be(
                "DatePicker「DeliveryDate」の日時を取得または設定します。");
    }
}
