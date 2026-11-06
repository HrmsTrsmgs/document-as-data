using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Marimo.DocumentAsData.Test.テスト補助;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class CheckBoxのテスト
{
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");
    static readonly string CheckedCheckBoxPath =
        Path.Combine("TestData", "チェック済みのチェックボックス.docx");
    static readonly string BoundCheckBoxPath =
        Path.Combine("TestData", "Custom XMLに連結されたチェックボックス.docx");

    [Fact]
    public void Documentプロパティはチェックボックスが属する文書を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.CheckBoxes.Single().Document.Should().BeSameAs(document);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void TagプロパティはチェックボックスのTagを取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.CheckBoxes.Single().Tag.Should().Be("Agreement");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ToStringは種類と引用符で囲んだTagを返します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.CheckBoxes.Single().ToString().Should().Be(
                "CheckBox { Tag = \"Agreement\" }");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void IsCheckedプロパティはチェック状態を取得します()
    {
        var uncheckedFilePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        var checkedFilePath =
            TestDocument.CreateTemporaryCopy(CheckedCheckBoxPath);
        try
        {
            using var uncheckedDocument = Document.Open(uncheckedFilePath, true);
            using var checkedDocument = Document.Open(checkedFilePath, true);

            uncheckedDocument.CheckBoxes.Single().IsChecked.Should().BeFalse();
            checkedDocument.CheckBoxes.Single().IsChecked.Should().BeTrue();
        }
        finally
        {
            File.Delete(uncheckedFilePath);
            File.Delete(checkedFilePath);
        }
    }

    [Fact]
    public void チェック状態要素が省略されたチェックボックスは未チェックです()
    {
        // w14:checkbox内のw14:checkedを省略し、表示文字には未チェックの☐を保持しています。
        var filePath = Path.Combine("TestData", "チェック状態要素を省略したチェックボックス.docx");
        using (var source = WordprocessingDocument.Open(filePath, false))
        {
            new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(source, TestCancellation.Token).Should().BeEmpty();
        }

        using var document = Document.Open(filePath, true);

        document.CheckBoxes["Agreement"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void チェック状態要素の値属性が省略されたチェックボックスはチェック済みです()
    {
        // w14:checkedは存在し、w14:valだけを省略しています。表示文字はチェック済みの☒です。
        var filePath = Path.Combine("TestData", "チェック状態の値属性を省略したチェックボックス.docx");
        using (var source = WordprocessingDocument.Open(filePath, false))
        {
            new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(source, TestCancellation.Token).Should().BeEmpty();
        }

        using var document = Document.Open(filePath, true);

        document.CheckBoxes["Agreement"].IsChecked.Should().BeTrue();
    }

    [Fact]
    public void CustomXMLに連結されたチェックボックスの状態は読み取りません()
    {
        // w:sdtPrのw:dataBindingがCustom XMLを指すため、表示側の状態だけを信用しません。
        using var document = Document.Open(BoundCheckBoxPath, true);
        var tested = document.CheckBoxes["Agreement"];

        var action = () => _ = tested.IsChecked;

        action.Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void IsCheckedプロパティはチェック状態を設定します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var tested = document.CheckBoxes.Single();

            tested.IsChecked = true;
            tested.IsChecked.Should().BeTrue();

            tested.IsChecked = false;
            tested.IsChecked.Should().BeFalse();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void チェック済み表示記号の値属性が省略されていてもチェック状態と既定の表示文字を設定します()
    {
        // 既存の未チェック文書からw14:checkedStateのw14:valだけを省略した固定DOCXです。
        // CheckedStateの仕様ではvalは省略可能で、チェック済みの既定文字はU+2612です。
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.checkedstate
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "チェック済み表示記号の値属性を省略したチェックボックス.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var source = WordprocessingDocument.Open(sourcePath, false))
            {
                new OpenXmlValidator(FileFormatVersions.Office2010)
                    .Validate(source, TestCancellation.Token).Should().BeEmpty();
            }

            using (var document = Document.Open(sourcePath, true))
            {
                document.CheckBoxes["Agreement"].IsChecked = true;

                document.CheckBoxes["Agreement"].IsChecked.Should().BeTrue();
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 状態のw14:checkedと表示文字のw:tは別なので、保存後の文字も確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("☒");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void CustomXMLに連結されたチェックボックスの状態は書き込みません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(BoundCheckBoxPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var tested = document.CheckBoxes["Agreement"];

            var action = () => tested.IsChecked = true;

            action.Should().Throw<NotSupportedException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void IsCheckedプロパティはチェック済みの表示文字を設定します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, true))
            {
                document.CheckBoxes.Single().IsChecked = true;
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // Wordのチェック状態はw14:checked、表示文字はw:tに別々に保持されます。
            // この文書のw14:checkedStateに指定されたU+2612をw:tへ反映し、
            // Wordで開いたときにもチェック済みとして表示されることを確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("☒");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }
}
