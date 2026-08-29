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
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields
                .Select(it => it.Name)
                .Should().Equal("CustomerName", "Address");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは名前からMERGEFIELDを取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Name.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは列挙と名前検索で同じMERGEFIELDを返します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);
            var enumerated = (
                from field in document.MergeFields
                where field.Name == "CustomerName"
                select field
            ).Single();

            document.MergeFields["CustomerName"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは存在しない名前を指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.MergeFields["not_found"];

            action.Should().Throw<KeyNotFoundException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
