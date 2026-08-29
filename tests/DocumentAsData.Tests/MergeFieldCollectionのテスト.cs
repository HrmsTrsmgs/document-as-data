using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldCollectionのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "simple-merge-fields.docx");

    [Fact]
    public void MergeFieldsは文書内のMERGEFIELDを列挙します()
    {
        using var document = Document.Open(TestFilePath);

        document.MergeFields
            .Select(it => it.Name)
            .Should().Equal("CustomerName", "Address");
    }

    [Fact]
    public void MergeFieldsは名前からMERGEFIELDを取得します()
    {
        using var document = Document.Open(TestFilePath);

        document.MergeFields["CustomerName"].Name.Should().Be("CustomerName");
    }

    [Fact]
    public void MergeFieldsは列挙と名前検索で同じMERGEFIELDを返します()
    {
        using var document = Document.Open(TestFilePath);
        var enumerated = (
            from field in document.MergeFields
            where field.Name == "CustomerName"
            select field
        ).Single();

        document.MergeFields["CustomerName"].Should().BeSameAs(enumerated);
    }

    [Fact]
    public void MergeFieldsは存在しない名前を指定した場合に失敗します()
    {
        using var document = Document.Open(TestFilePath);

        var action = () => _ = document.MergeFields["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }
}
