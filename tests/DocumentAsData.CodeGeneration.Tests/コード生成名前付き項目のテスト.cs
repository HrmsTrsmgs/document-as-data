using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前付き項目のテスト : IDisposable
{
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";

    readonly TemporaryDocumentFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

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

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           MergeFieldsDocumentFilePath))
                   .GeneratedType("MERGEFIELDDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       MergeFieldsDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.MergeFields["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           TextContentControlDocumentFilePath))
                   .GeneratedType("文字列ContentControlDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       TextContentControlDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.ContentControls["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 生成されたDocument型のCheckBoxプロパティからチェック状態を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.Agreement;

        tested.Should().BeTrue();
    }
}
