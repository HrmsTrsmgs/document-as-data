using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldCollectionのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "単純形式のMERGEFIELD.docx");
    static readonly string DuplicateMergeFieldsPath =
        Path.Combine("TestData", "同名のMERGEFIELD.docx");
    static readonly string StandardComplexMergeFieldPath =
        Path.Combine(
            "TestData",
            "Word標準の書式維持スイッチを持つ複合MERGEFIELD.docx");
    static readonly string MultipleComplexMergeFieldsPath =
        Path.Combine("TestData", "複数の複合MERGEFIELD.docx");
    static readonly string OtherFieldsPath =
        Path.Combine("TestData", "MERGEFIELD以外のフィールド.docx");

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

    [Fact]
    public void MergeFieldsは同名のMERGEFIELDが複数存在する場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.MergeFields["CustomerName"];

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "Word標準の書式維持スイッチを持つ固定テストデータを追加してから有効化します。")]
    public void MergeFieldsは書式維持スイッチを含むWord標準の命令からMERGEFIELD名を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(StandardComplexMergeFieldPath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields.Single().Name.Should().Be("CustomerName");
            document.MergeFields["CustomerName"].Should().NotBeNull();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "複数の複合MERGEFIELDを持つ固定テストデータを追加してから有効化します。")]
    public void MergeFieldsは複数の複合MERGEFIELDを文書内の順序で列挙します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(MultipleComplexMergeFieldsPath);
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

    [Fact(Skip = "MERGEFIELD以外のWordフィールドを持つ固定テストデータを追加してから有効化します。")]
    public void MergeFieldsはMERGEFIELD以外のWordフィールドを列挙しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(OtherFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields.Should().BeEmpty();
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
