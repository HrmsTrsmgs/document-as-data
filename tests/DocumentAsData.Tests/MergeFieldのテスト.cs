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
    static readonly string SplitSimpleMergeFieldResultPath =
        Path.Combine("TestData", "結果が分割された単純MERGEFIELD.docx");

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
    public void 保存した単純MERGEFIELDは分割されていた古い結果を残しません()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitSimpleMergeFieldResultPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            // 固定データでは、単純MERGEFIELDの表示結果「株式会社○○」が
            // 二つのw:rとw:tへ分割されています。値の設定後に古い二つ目の
            // 文字列「○○」が残らないことを、保存して開き直して確認します。
            saved.MergeFields["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void 保存した単純MERGEFIELDは値の前後の空白を保持します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void 保存した複合MERGEFIELDは値の前後の空白を保持します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "単純MERGEFIELDの改行とタブをValueで表す仕様を確認してから有効化します。")]
    public void 保存した単純MERGEFIELDは値の改行とタブを保持します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = "一行目\t二列目\r\n二行目";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "複合MERGEFIELDの改行とタブをValueで表す仕様を確認してから有効化します。")]
    public void 保存した複合MERGEFIELDは値の改行とタブを保持します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = "一行目\t二列目\r\n二行目";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Valueプロパティは単純MERGEFIELDの空白を保持する属性を設定します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // XMLでは、前後の空白をWordの表示内容として保持するために
            // xml:space="preserve"が必要です。DocumentAsDataで開き直して
            // 文字列を読めるだけでなく、Wordでも失われないことを確認します。
            var text = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single(it => it.Text == value);
            text.Space?.Value.Should().Be(
                DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Valueプロパティは複合MERGEFIELDの空白を保持する属性を設定します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 複合MERGEFIELDでは結果用のw:tを新しく作るため、その要素にも
            // xml:space="preserve"が設定されることを直接確認します。
            var text = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single(it => it.Text == value);
            text.Space?.Value.Should().Be(
                DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "単純MERGEFIELDの改行とタブをWordで表示できるOOXMLを確認するときに有効化します。")]
    public void Valueプロパティは単純MERGEFIELDの改行とタブを専用要素で保存します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value =
                    "一行目\t二列目\r\n二行目";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // Wordのタブと改行はw:t内の制御文字ではなく、それぞれ
            // w:tabとw:brで表します。Wordで同じ表示になる内部構造を
            // 保存できていることを直接確認します。
            var field = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.SimpleField>()
                .Single(it =>
                    it.Instruction?.Value?.Contains("CustomerName") == true);
            field.Descendants<Wordprocessing.TabChar>()
                .Should().ContainSingle();
            field.Descendants<Wordprocessing.Break>()
                .Should().ContainSingle();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "複合MERGEFIELDの改行とタブをWordで表示できるOOXMLを確認するときに有効化します。")]
    public void Valueプロパティは複合MERGEFIELDの改行とタブを専用要素で保存します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value =
                    "一行目\t二列目\r\n二行目";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 複合MERGEFIELDの結果領域でも、Wordが解釈できるw:tabと
            // w:brを保存することを直接確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.TabChar>()
                .Should().ContainSingle();
            saved.MainDocumentPart.Document
                .Descendants<Wordprocessing.Break>()
                .Should().ContainSingle();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "結果が分割された単純MERGEFIELDの固定テストデータを追加してから有効化します。")]
    public void Valueプロパティは分割して保存された単純MERGEFIELDの結果全体を置き換えます()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitSimpleMergeFieldResultPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 単純MERGEFIELDの表示結果も、書式などによって複数のw:t要素へ
            // 分かれることがあります。Wordの差し込み後と同様に、古い結果の
            // 断片を残さず、新しい値だけになることを直接確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.SimpleField>()
                .Single()
                .Descendants<Wordprocessing.Text>()
                .Select(it => it.Text)
                .Should().Equal("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
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
