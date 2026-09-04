using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class CheckBoxCollectionのテスト
{
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");

    [Fact]
    public void CheckBoxesは文書内のチェックボックスを列挙します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.CheckBoxes.Should().ContainSingle();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CheckBoxesはTagからチェックボックスを取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.CheckBoxes["Agreement"].Tag.Should().Be("Agreement");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CheckBoxesは列挙とTag検索で同じチェックボックスを返します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(CheckBoxContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var enumerated = document.CheckBoxes.Single();

            document.CheckBoxes["Agreement"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
