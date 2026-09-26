using FluentAssertions;

namespace Marimo.DocumentAsData.Test;

public class RepeatingSectionCollectionのテスト
{
    [Fact(Skip = "繰り返しセクションのAPIレビュー後、このテストからRedを再開します。")]
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

    [Fact(Skip = "セクション列挙のGreen後、文書順と別セクションの識別をレビューします。")]
    public void RepeatingSectionsは複数のセクションを文書順に列挙します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "異なるTagの繰り返しセクション.docx"));

        document.RepeatingSections.Select(it => it.Tag).Should().Equal("Items", "Options");
    }

    [Fact(Skip = "セクション列挙のGreen後、通常の入力項目との区別をレビューします。")]
    public void RepeatingSectionsは通常のContentControlを含めません()
    {
        using var document = Document.Open(Path.Combine("TestData", "単一のContent Control.docx"));

        document.RepeatingSections.Should().BeEmpty();
    }

    [Fact(Skip = "セクション列挙のGreen後、Tagによる取得をレビューします。")]
    public void RepeatingSectionsはTagでセクションを取得します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "異なるTagの繰り返しセクション.docx"));

        document.RepeatingSections["Options"].Tag.Should().Be("Options");
    }

    [Fact(Skip = "Tagによる取得のGreen後、存在しないセクションの扱いをレビューします。")]
    public void RepeatingSectionsは存在しないTagで失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションのContent Control.docx"));

        var action = () => document.RepeatingSections["Missing"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact(Skip = "Tagによる取得のGreen後、セクション自体の同名重複をレビューします。")]
    public void RepeatingSectionsは同じTagのセクションが複数あると失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "同じTagの繰り返しセクション.docx"));

        var action = () => document.RepeatingSections["Items"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact(Skip = "セクション列挙のGreen後、Tagなしの枠を名前付き項目から除く仕様をレビューします。")]
    public void RepeatingSectionsはTagのないセクションを列挙しません()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "Tagなしを含む繰り返しセクション.docx"));

        document.RepeatingSections.Select(it => it.Tag).Should().Equal("Items");
    }
}
