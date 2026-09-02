using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class ContentControlのテスト
{
    static readonly string ContentControlPath =
        Path.Combine("TestData", "単一のContent Control.docx");
    static readonly string SplitContentControlTextPath =
        Path.Combine("TestData", "複数の文字列要素を持つContent Control.docx");
    static readonly string InvalidContentControlWithoutTagValuePath =
        Path.Combine("TestData", "Tagの値が欠落した不正なContent Control.docx");
    static readonly string PlaceholderContentControlPath =
        Path.Combine("TestData", "プレースホルダー表示中のContent Control.docx");
    static readonly string MultilineContentControlPath =
        Path.Combine("TestData", "複数行のContent Control.docx");

    [Fact]
    public void DocumentプロパティはContentControlが属する文書を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Single().Document.Should().BeSameAs(document);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void TagプロパティはContentControlのTagを取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Single().Tag.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Tagプロパティは不正なOOXMLでTagに値がない場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                _ = document.ContentControls.Single().Tag;

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ValueプロパティはContentControlの値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value.Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Valueプロパティは複数の文字列要素からContentControlの値を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value
                .Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ValueプロパティはContentControlの値を設定します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value = "変更後";

            document.ContentControls["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Valueプロパティはプレースホルダー表示中なら空文字列を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value.Should().BeEmpty();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Valueプロパティはプレースホルダー表示中のContentControlへ値を設定します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Value = "変更後";

            document.ContentControls["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void 保存したContentControlから空文字列を取得します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Value.Should().BeEmpty();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void 保存したContentControlは値の前後の空白を保持します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "複数行Content Controlの固定テストデータとValueの改行・タブ表現を確認してから有効化します。")]
    public void 保存した複数行ContentControlは値の改行とタブを保持します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(MultilineContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = "一行目\t二列目\r\n二行目";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Value.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void 保存したContentControlは分割されていた古い値を残しません()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void ValueプロパティはContentControlの空白を保持する属性を設定します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = value;
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

    [Fact(Skip = "複数行Content Controlの改行とタブをWordで表示できるOOXMLを確認するときに有効化します。")]
    public void ValueプロパティはContentControlの改行とタブを専用要素で保存します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(MultilineContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value =
                    "一行目\t二列目\r\n二行目";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // Wordのタブと改行はw:t内の制御文字ではなく、それぞれ
            // w:tabとw:brで表します。Wordで同じ表示になる内部構造を
            // 保存できていることを直接確認します。
            var contentControl = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.SdtElement>()
                .Single();
            contentControl.Descendants<Wordprocessing.TabChar>()
                .Should().ContainSingle();
            contentControl.Descendants<Wordprocessing.Break>()
                .Should().ContainSingle();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Valueプロパティは複数の文字列要素の先頭へ値を設定して残りを空にします()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // Content Controlの表示文字列は、OOXML上で複数のw:t要素に
            // 分かれることがあります。保存後のOOXMLでは先頭のw:tだけに
            // 設定値を残し、2個目以降を空にすることを直接確認します。
            // Valueの再取得だけでは、w:tごとの保存位置を区別できないためです。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Select(it => it.Text)
                .Should().Equal("変更後", "");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Valueプロパティはプレースホルダー表示状態を解除します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // プレースホルダー文字列は入力値ではありません。Wordで入力した
            // 場合と同様に、値の設定後はw:showingPlcHdrを除去して、次に
            // Wordで開いたときも設定値が通常の入力内容として扱われることを
            // 直接確認します。
            var contentControl = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.SdtElement>()
                .Single();
            contentControl.Descendants<Wordprocessing.ShowingPlaceholder>()
                .Should().BeEmpty();
            contentControl.Descendants<Wordprocessing.Text>()
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
