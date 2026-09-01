using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "単純形式のMERGEFIELD.docx");
    static readonly string ComplexMergeFieldPath =
        Path.Combine("TestData", "複合形式のMERGEFIELD.docx");
    static readonly string SplitComplexMergeFieldPath =
        Path.Combine("TestData", "命令が分割された複合MERGEFIELD.docx");
    static readonly string SplitComplexMergeFieldResultPath =
        Path.Combine("TestData", "結果が分割された複合MERGEFIELD.docx");

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

    [Fact]
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

    [Fact]
    public void Valueプロパティは分割して保存された複合MERGEFIELDの結果全体を置き換えます()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitComplexMergeFieldResultPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // MERGEFIELDの表示結果は、書式などによって複数のw:t要素へ
            // 分かれることがあります。Wordの差し込みと同様に、設定後は
            // 古い結果を残さず、新しい値だけになることを直接確認します。
            var fieldCharacters = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.FieldChar>();
            var resultStart = fieldCharacters.Single(
                it => it.FieldCharType?.Value ==
                    Wordprocessing.FieldCharValues.Separate).Parent;
            var resultEnd = fieldCharacters.Single(
                it => it.FieldCharType?.Value ==
                    Wordprocessing.FieldCharValues.End).Parent;
            resultStart.Should().NotBeNull();
            resultEnd.Should().NotBeNull();

            var resultElements = resultStart.ElementsAfter()
                .TakeWhile(it => it != resultEnd)
                .ToArray();
            resultElements.Should().ContainSingle();
            var resultRun = resultElements.Single()
                .Should().BeOfType<Wordprocessing.Run>().Which;
            resultRun.ChildElements.Should().ContainSingle();
            resultRun.Elements<Wordprocessing.Text>()
                .Select(it => it.Text)
                .Should().Equal("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

}
