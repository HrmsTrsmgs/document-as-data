using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class ContentControlのテスト
{
    static readonly string ContentControlPath =
        Path.Combine("TestData", "content-control.docx");

    [Fact(Skip = "Content ControlとDocumentの関連をGreen対象にするときに有効化します。")]
    public void DocumentプロパティはContentControlが属する文書を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Single().Document.Should().BeSameAs(document);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "Content ControlのTag取得をGreen対象にするときに有効化します。")]
    public void TagプロパティはContentControlのTagを取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Single().Tag.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "Content Controlの値取得をGreen対象にするときに有効化します。")]
    public void ValueプロパティはContentControlの値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value.Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "Content Controlの値設定をGreen対象にするときに有効化します。")]
    public void ValueプロパティはContentControlの値を設定します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value = "変更後";

            document.ContentControls["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

}
