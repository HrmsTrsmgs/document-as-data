using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class TextContentControlのテスト
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
    static readonly string BoundContentControlPath =
        Path.Combine("TestData", "Custom XMLに連結された文字列Content Control.docx");

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
    public void ToStringは種類と引用符で囲んだTagを返します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls.Single().ToString().Should().Be(
                "TextContentControl { Tag = \"CustomerName\" }");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ToStringはTagのバックスラッシュに続く引用符をそれぞれエスケープします()
    {
        using var document = Document.Open(@"TestData\引用符とバックスラッシュを含むTag.docx");

        document.ContentControls.Single().ToString().Should().Be(
            """TextContentControl { Tag = "C:\\temp\\\"document" }""");
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
    public void TextプロパティはContentControlの値を取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Text.Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CustomXMLに連結されたContentControlの文字列は読み取りません()
    {
        // w:sdtPrのw:dataBindingが、Custom XMLの/root/valueを指します。
        // w:sdtContentの表示文字だけでは、連結先の現在値を保証できません。
        using var document = Document.Open(BoundContentControlPath, true);
        var tested = document.ContentControls["CustomerName"];

        var action = () => _ = tested.Text;

        action.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void CustomXMLに連結されたContentControlの文字列は書き込みません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(BoundContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var tested = document.ContentControls["CustomerName"];

            var action = () => tested.Text = "変更後";

            action.Should().Throw<NotSupportedException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティは複数の文字列要素からContentControlの値を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Text
                .Should().Be("山田太郎");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void TextプロパティはContentControl内の段落を改行で区切って取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "二つの段落を持つContent Control.docx"));
        try
        {
            // 一つのw:sdtContent内に「前半」と「後半」のw:pが一つずつあります。
            // w:pは段落、w:rは文字のまとまり、w:tは文字列です。
            // w:brによる段落内改行ではなく、段落の境界をCRLFとして読み取ります。
            // 固定DOCXはこの構造を狙った最小OOXMLです。
            using var document = Document.Open(filePath, validate: true);

            document.ContentControls["CustomerName"].Text.Should().Be("前半\r\n後半");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティは複数段落から読み取った値を再設定しても改行を増やしません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "二つの段落を持つContent Control.docx"));
        try
        {
            // w:sdtContent内に「前半」と「後半」のw:p（段落）が一つずつあります。
            // 読み取りでは段落境界がCRLFになります。書き込みで後続段落を空にするだけだと、
            // 再読み取り時にその空段落の分だけ末尾のCRLFが増えてしまいます。
            using var document = Document.Open(filePath, validate: true);
            var tested = document.ContentControls["CustomerName"];

            tested.Text = tested.Text;

            tested.Text.Should().Be("前半\r\n後半");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティは後続段落にある画像を保存後も保持します()
    {
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(
                @"TestData\後続段落に画像を持つ文字列Content Control.docx", validate: true))
            {
                document.ContentControls["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 二つ目のw:pには文字列ではなくw:drawing（画像）があります。
            // 段落ごと消すと画像への参照も失われるため、保存された要素を確認します。
            saved.MainDocumentPart.Should().NotBeNull();
            saved.MainDocumentPart.Document.Should().NotBeNull();
            saved.MainDocumentPart.Document.Descendants<Wordprocessing.Drawing>()
                .Should().ContainSingle();
            new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2010)
                .Validate(saved).Should().BeEmpty();
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Textプロパティは外側を書き換えても内側のContentControlを引き続き編集できます()
    {
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(
                @"TestData\後続段落が別のContent Controlに属する文書.docx", validate: true))
            {
                // Outerのw:sdtContentには直接のw:pと、Innerのw:sdtがあります。
                // Inner内部のw:pはOuterの後続段落として消してはいけません。
                document.ContentControls["Outer"].Text = "変更後";
                document.ContentControls["Inner"].Text = "内側の変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath, validate: true);
            saved.ContentControls.Select(it => it.Tag).Should().Equal("Outer", "Inner");
            saved.ContentControls["Inner"].Text.Should().Be("内側の変更後");
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void TextプロパティはContentControlの値を設定します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Text = "変更後";

            document.ContentControls["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティはプレースホルダー表示中なら空文字列を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Text.Should().BeEmpty();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティはプレースホルダー表示中のContentControlへ値を設定します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        try
        {
            using var document = Document.Open(filePath);

            document.ContentControls["CustomerName"].Text = "変更後";

            document.ContentControls["CustomerName"].Text.Should().Be("変更後");
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
                document.ContentControls["CustomerName"].Text = "";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Text.Should().BeEmpty();
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
                document.ContentControls["CustomerName"].Text = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Text.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
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
                document.ContentControls["CustomerName"].Text = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Text.Should().Be(value);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Textプロパティは改行とタブを含むContentControlへ再設定したとき最後の値だけを取得します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);
            var tested = document.ContentControls["CustomerName"];
            tested.Text = "一行目\t二列目\r\n二行目";

            tested.Text = "再設定";

            tested.Text.Should().Be("再設定");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Textプロパティは空文字列を設定したContentControlへ値を再設定します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);
            var tested = document.ContentControls["CustomerName"];
            tested.Text = "";

            tested.Text = "再設定";

            tested.Text.Should().Be("再設定");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Theory]
    [InlineData("\t")]
    [InlineData("\r\n")]
    [InlineData("\n")]
    public void Textプロパティはタブまたは改行だけを設定したContentControlへ値を再設定します(string value)
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        try
        {
            using var document = Document.Open(filePath);
            var tested = document.ContentControls["CustomerName"];
            // タブと改行はOOXMLではw:tではなくw:tabとw:brになります。
            // それらだけの値を書いた後も、次の文字列を書き込める必要があります。
            tested.Text = value;

            tested.Text = "再設定";

            tested.Text.Should().Be("再設定");
        }
        finally
        {
            File.Delete(filePath);
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
                document.ContentControls["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void TextプロパティはContentControlの空白を保持する属性を設定します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        const string value = " 変更後 ";
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text = value;
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
    public void TextプロパティはContentControlの改行とタブを専用要素で保存します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(MultilineContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text =
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
    public void TextプロパティはLF改行をContentControlの改行要素で保存します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text =
                    "一行目\n二行目";
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // .NETの文字列ではLF単独も改行として使われますが、Word文書では
            // w:t内のLFではなくw:br要素として保存する必要があります。
            var contentControl = saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.SdtElement>()
                .Single();
            contentControl.Descendants<Wordprocessing.Break>()
                .Should().ContainSingle();
            contentControl.Descendants<Wordprocessing.Text>()
                .Select(it => it.Text)
                .Should().Equal("一行目", "二行目");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Textプロパティは複数の文字列要素の先頭へ値を設定して残りを空にします()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(SplitContentControlTextPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text = "変更後";
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
    public void Textプロパティはプレースホルダー表示状態を解除します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(PlaceholderContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text = "変更後";
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
