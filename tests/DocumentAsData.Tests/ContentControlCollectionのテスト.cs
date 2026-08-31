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
    static readonly string TableContentControlPath =
        Path.Combine("TestData", "表内のContent Control.docx");
    static readonly string NestedContentControlsPath =
        Path.Combine("TestData", "ネストしたContent Control.docx");
    static readonly string DuplicateContentControlsPath =
        Path.Combine("TestData", "同じTagのContent Control.docx");

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

    [Fact(Skip = "Content ControlのTag検索をGreen対象にするときに有効化します。")]
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

    [Fact(Skip = "Content Controlオブジェクトの同一性をGreen対象にするときに有効化します。")]
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

    [Fact(Skip = "存在しないContent Control Tagの扱いをGreen対象にするときに有効化します。")]
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

    [Fact(Skip = "同一TagのContent Control重複検索をGreen対象にするときに有効化します。")]
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

    [Fact(Skip = "TagなしContent Controlの除外をGreen対象にするときに有効化します。")]
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

    [Fact(Skip = "表内のContent Control列挙をGreen対象にするときに有効化します。")]
    public void ContentControlsは表内のContentControlも列挙します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TableContentControlPath);
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

    [Fact(Skip = "ネストしたContent Control列挙をGreen対象にするときに有効化します。")]
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
}
