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

    [Fact(Skip = "MERGEFIELDオブジェクトの同一性をGreen対象にするときに有効化します。")]
    public void MergeFieldsは列挙と名前検索で同じMERGEFIELDを返します()
    {
        var filePath = TestDocument.CreateWithSimpleMergeFields(("CustomerName", "株式会社○○"));
        try
        {
            using var document = Document.Open(filePath);
            var enumerated = document.MergeFields.Single();

            document.MergeFields["CustomerName"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "存在しないMERGEFIELD名の扱いをGreen対象にするときに有効化します。")]
    public void MergeFieldsは存在しない名前を指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateCopy();
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

    [Fact(Skip = "表内のMERGEFIELD列挙をGreen対象にするときに有効化します。")]
    public void MergeFieldsは表内のMERGEFIELDも列挙します()
    {
        var filePath = TestDocument.CreateWithTableMergeField("CustomerName", "株式会社○○");
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields.Single().Name.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
