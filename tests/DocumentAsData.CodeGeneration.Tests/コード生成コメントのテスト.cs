using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

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
    public void 文書名とTagにアンパサンドを含んでも正しいXMLコメントを生成します()
    {
        // 文書名もTagもR&Dです。DOCX内のTagはXML上でR&amp;Dと保存しています。
        // C#の文字列に埋め込めても、XMLコメント内の&は別途エスケープが必要です。
        // 通常のコンパイル補助はコメント警告を検出しないため、ここでは明示的に診断します。
        var sources = GeneratedCodeInspection.GenerateSources(
            @"TestData\コード生成\R&D.docx",
            options => options.NameMappings["R&D"] = "Research");

        (
            from source in sources
            from diagnostic in
                CSharpSyntaxTree.ParseText(
                    source,
                    new CSharpParseOptions(documentationMode: DocumentationMode.Diagnose))
                    .GetDiagnostics()
            where diagnostic.Severity is DiagnosticSeverity.Warning or DiagnosticSeverity.Error
            select diagnostic
        ).Should().BeEmpty();

        sources.TypeDeclaration("ResearchDocument")
            .SummaryText().Should().Be("Word文書「R&D」を型付きで表します。");
        sources.TypeDeclaration("ResearchDocument")
            .PropertyDeclaration("Research")
            .SummaryText().Should().Be("文字列Content Control「R&D」の文字列を取得または設定します。");
    }

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
            .MethodDeclaration("Open", "Stream")
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
    public void Data型の文字列プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(MergeFieldsDocumentFilePath)
            .TypeDeclaration("MergefieldData")
            .PropertyDeclaration("CustomerName")
            .SummaryText()
            .Should().Be(
                "文書項目「CustomerName」の文字列を取得または設定します。");
    }

    [Fact]
    public void Data型のCheckBoxプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(CheckBoxDocumentFilePath)
            .TypeDeclaration("チェック済みCheckBoxData")
            .PropertyDeclaration("Agreement")
            .SummaryText()
            .Should().Be(
                "CheckBox「Agreement」のチェック状態を取得または設定します。");
    }

    [Fact]
    public void Data型のDatePickerプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DatePickerDocumentFilePath)
            .TypeDeclaration("日付選択ContentControlData")
            .PropertyDeclaration("DeliveryDate")
            .SummaryText()
            .Should().Be(
                "DatePicker「DeliveryDate」の日時を取得または設定します。");
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
