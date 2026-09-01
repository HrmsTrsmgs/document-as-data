using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class ContentControlのテスト
{
    static readonly string ContentControlPath =
        Path.Combine("TestData", "単一のContent Control.docx");
    static readonly string SplitContentControlTextPath =
        Path.Combine("TestData", "複数の文字列要素を持つContent Control.docx");
    static readonly string InvalidContentControlWithoutTagValuePath =
        Path.Combine("TestData", "Tagの値が欠落した不正なContent Control.docx");

    [Fact]
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

    [Fact]
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

    [Fact]
    public void Tagプロパティは不正なOOXMLでTagに値がない場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                _ = document.ContentControls.Single().Tag;

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
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

    [Fact]
    public void Valueプロパティは複数の文字列要素からContentControlの値を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value
                .Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
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

    [Fact]
    public void Valueプロパティは複数の文字列要素の先頭へ値を設定して残りを空にします()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // Content Controlの表示文字列は、OOXML上で複数のw:t要素に
            // 分かれることがあります。保存後のOOXMLでは先頭のw:tだけに
            // 設定値を残し、2個目以降を空にすることを直接確認します。
            // Valueの再取得だけでは、w:tごとの保存位置を区別できないためです。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Select(it => it.Text)
                .Should().Equal("変更後", "");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

}
