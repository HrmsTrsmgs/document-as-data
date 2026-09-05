using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class CheckBoxのテスト
{
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");
    static readonly string CheckedCheckBoxPath =
        Path.Combine("TestData", "チェック済みのチェックボックス.docx");

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
