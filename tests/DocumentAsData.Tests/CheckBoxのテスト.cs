using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class CheckBoxのテスト
{
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");
    static readonly string CheckedCheckBoxPath =
        Path.Combine("TestData", "チェック済みのチェックボックス.docx");

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
}
