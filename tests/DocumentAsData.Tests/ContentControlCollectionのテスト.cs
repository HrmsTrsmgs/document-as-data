using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class ContentControlCollectionのテスト
{
    static readonly string EmptyDocumentPath =
        Path.Combine("TestData", "空の文書.docx");
    static readonly string ContentControlPath =
        Path.Combine("TestData", "単一のContent Control.docx");
    static readonly string ContentControlsPath =
        Path.Combine("TestData", "複数のContent Control.docx");
    static readonly string ContentControlsWithoutTagPath =
        Path.Combine("TestData", "Tagなしを含むContent Control.docx");
    static readonly string NestedContentControlsPath =
        Path.Combine("TestData", "ネストしたContent Control.docx");
    static readonly string DuplicateContentControlsPath =
        Path.Combine("TestData", "同じTagのContent Control.docx");
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");

    [Fact]
    public void ContentControlsは文書内のContentControlを列挙します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Should().HaveCount(2);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsは文書内の順序で列挙します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls
                .Select(it => it.Tag)
                .Should().Equal("CustomerName", "Address");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsはTagからContentControlを取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Tag.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsは列挙とTag検索で同じContentControlを返します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);
            var enumerated = document.ContentControls.Single();

            document.ContentControls["CustomerName"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsは存在しないTagを指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.ContentControls["not_found"];

            action.Should().Throw<KeyNotFoundException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsは同じTagのContentControlが複数存在する場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.ContentControls["CustomerName"];

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsはTagのないContentControlを列挙しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsWithoutTagPath);
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
    public void ContentControlsはネストしたContentControlをそれぞれ列挙します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(NestedContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls
                .Select(it => it.Tag)
                .Should().Equal("Outer", "Inner");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ContentControlsはチェックボックスを列挙しません()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            // Wordのチェックボックスは、表示記号のw:tとは別に
            // w:sdtPr/w14:checkbox/w14:checkedへチェック状態を保持します。
            // 文字列用のValueでは両者を同期できないため、通常のContentControlsには含めません。
            document.ContentControls.Should().BeEmpty();
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
