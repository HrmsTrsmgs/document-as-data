using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "simple-merge-fields.docx");

    [Fact]
    public void DocumentプロパティはMERGEFIELDが属する文書を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Document.Should().BeSameAs(document);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void NameプロパティはMERGEFIELD名を取得します()
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

    [Fact(Skip = "simple MERGEFIELDの値取得をGreen対象にするときに有効化します。")]
    public void Valueプロパティはsimple形式のMERGEFIELD値を取得します()
    {
        var filePath = TestDocument.CreateWithSimpleMergeFields(("CustomerName", "株式会社○○"));
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Value.Should().Be("株式会社○○");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "複合MERGEFIELDの読み取りをGreen対象にするときに有効化します。")]
    public void 複合MERGEFIELDの名前と値を取得します()
    {
        var filePath = TestDocument.CreateWithComplexMergeField("CustomerName", "株式会社○○");
        try
        {
            using var document = Document.Open(filePath);
            var tested = document.MergeFields.Single();

            tested.Name.Should().Be("CustomerName");
            tested.Value.Should().Be("株式会社○○");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "分割された複合MERGEFIELD命令の読み取りをGreen対象にするときに有効化します。")]
    public void 複数のinstrTextに分割された複合MERGEFIELDを取得します()
    {
        var filePath = TestDocument.CreateWithSplitComplexMergeField("CustomerName", "株式会社○○");
        try
        {
            using var document = Document.Open(filePath);
            var tested = document.MergeFields.Single();

            tested.Name.Should().Be("CustomerName");
            tested.Value.Should().Be("株式会社○○");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "simple MERGEFIELDの値設定をGreen対象にするときに有効化します。")]
    public void Valueプロパティはsimple形式のMERGEFIELD値を設定します()
    {
        var filePath = TestDocument.CreateWithSimpleMergeFields(("CustomerName", "変更前"));
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Value = "変更後";

            document.MergeFields["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "複合MERGEFIELDの値設定をGreen対象にするときに有効化します。")]
    public void Valueプロパティは複合MERGEFIELDの値を設定します()
    {
        var filePath = TestDocument.CreateWithComplexMergeField("CustomerName", "変更前");
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Value = "変更後";

            document.MergeFields["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact(Skip = "同名MERGEFIELDの一括更新をGreen対象にするときに有効化します。")]
    public void Valueプロパティは同名のMERGEFIELDをすべて更新します()
    {
        var filePath = TestDocument.CreateWithSimpleMergeFields(
            ("CustomerName", "変更前1"),
            ("CustomerName", "変更前2"));
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields["CustomerName"].Value = "変更後";

            document.MergeFields
                .Where(it => it.Name == "CustomerName")
                .Select(it => it.Value)
                .Should().OnlyContain(it => it == "変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
