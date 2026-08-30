using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "simple-merge-fields.docx");
    static readonly string ComplexMergeFieldPath =
        Path.Combine("TestData", "complex-merge-field.docx");
    static readonly string SplitComplexMergeFieldPath =
        Path.Combine("TestData", "split-complex-merge-field.docx");
    static readonly string DuplicateMergeFieldsPath =
        Path.Combine("TestData", "duplicate-merge-fields.docx");

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

    [Fact]
    public void ValueプロパティはMERGEFIELD値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
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

    [Fact]
    public void 複合MERGEFIELDの名前と値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
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

    [Fact]
    public void 複合MERGEFIELDが分割して保存されていても名前と値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SplitComplexMergeFieldPath);
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

    [Fact]
    public void ValueプロパティはMERGEFIELD値を設定します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
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
        var filePath = TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
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
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateMergeFieldsPath);
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
