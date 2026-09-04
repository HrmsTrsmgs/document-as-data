using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class CheckBoxCollectionのテスト
{
    static readonly string EmptyDocumentPath =
        Path.Combine("TestData", "空の文書.docx");

    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");

    static readonly string CheckBoxesWithoutTagPath =
        Path.Combine("TestData", "Tagなしを含むチェックボックス.docx");

    static readonly string DuplicateCheckBoxesPath =
        Path.Combine("TestData", "同じTagのチェックボックス.docx");

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

    [Fact]
    public void CheckBoxesは存在しないTagを指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () => _ = document.CheckBoxes["not_found"];

            action.Should().Throw<KeyNotFoundException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CheckBoxesは同じTagのチェックボックスが複数存在する場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateCheckBoxesPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () => _ = document.CheckBoxes["Agreement"];

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CheckBoxesはTagのないチェックボックスを列挙しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(CheckBoxesWithoutTagPath);
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
}
