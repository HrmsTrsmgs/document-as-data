using FluentAssertions;

namespace Marimo.DocumentAsData.Test;

public class RepeatingSectionCollectionのテスト
{
    [Fact]
    public void RepeatingSectionsは明細が1件でも繰り返しセクションを列挙します()
    {
        // Itemsはw15:repeatingSectionで繰り返し全体を表し、その内側に
        // w15:repeatingSectionItemで示された明細Itemが1件だけあります。
        // 明細の中には文字列Content ControlのItemNameがありますが、
        // 繰り返しセクションとして取得するのは外側のItemsだけです。
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        document.RepeatingSections
            .Select(it => it.Tag)
            .Should().Equal("Items");
    }

    [Fact]
    public void RepeatingSectionsは複数のセクションを文書順に列挙します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "異なるTagの繰り返しセクション.docx"));

        document.RepeatingSections.Select(it => it.Tag).Should().Equal("Items", "Options");
    }

    [Fact]
    public void RepeatingSectionsは通常のContentControlを含めません()
    {
        using var document = Document.Open(Path.Combine("TestData", "単一のContent Control.docx"));

        document.RepeatingSections.Should().BeEmpty();
    }

    [Fact]
    public void RepeatingSectionsはTagでセクションを取得します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "異なるTagの繰り返しセクション.docx"));

        document.RepeatingSections["Options"].Tag.Should().Be("Options");
    }

    [Fact]
    public void RepeatingSectionsは存在しないTagで失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        var action = () => document.RepeatingSections["Missing"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void RepeatingSectionsは同じTagのセクションが複数あると失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "同じTagの繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void RepeatingSectionsはTagのないセクションを列挙しません()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "Tagなしを含む繰り返しセクション.docx"));

        document.RepeatingSections.Select(it => it.Tag).Should().Equal("Items");
    }
}
