using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class ContentControlCollectionのテスト
{
    [Fact(Skip = "Content Control列挙をGreen対象にするときに有効化します。")]
    public void ContentControlsは文書内のContentControlを列挙します()
    {
        var filePath = TestDocument.CreateWithContentControls(
            ("CustomerName", "山田太郎"),
            ("Address", "東京都"));
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
        var filePath = TestDocument.CreateWithContentControls(("CustomerName", "山田太郎"));
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
        var filePath = TestDocument.CreateWithContentControls(("CustomerName", "山田太郎"));
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
        var filePath = TestDocument.CreateCopy();
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

    [Fact(Skip = "TagなしContent Controlの除外をGreen対象にするときに有効化します。")]
    public void ContentControlsはTagのないContentControlを列挙しません()
    {
        var filePath = TestDocument.CreateWithContentControls(
            (null, "名前なし"),
            ("CustomerName", "山田太郎"));
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
        var filePath = TestDocument.CreateWithTableContentControl("CustomerName", "山田太郎");
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
        var filePath = TestDocument.CreateWithNestedContentControls("Outer", "Inner", "山田太郎");
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
