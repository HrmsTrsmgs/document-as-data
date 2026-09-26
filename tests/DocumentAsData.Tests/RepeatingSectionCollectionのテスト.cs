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
}
