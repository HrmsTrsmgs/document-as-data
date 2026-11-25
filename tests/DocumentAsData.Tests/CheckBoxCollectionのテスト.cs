using FluentAssertions;

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
        using var document = Document.Open(CheckBoxContentControlPath, validate: true);

        document.CheckBoxes.Should().ContainSingle();
    }

    [Fact]
    public void CheckBoxesはTagからチェックボックスを取得します()
    {
        using var document = Document.Open(CheckBoxContentControlPath, validate: true);

        document.CheckBoxes["Agreement"].Tag.Should().Be("Agreement");
    }

    [Fact]
    public void CheckBoxesは列挙とTag検索で同じチェックボックスを返します()
    {
        using var document = Document.Open(CheckBoxContentControlPath, validate: true);
        var enumerated = document.CheckBoxes.Single();

        document.CheckBoxes["Agreement"].Should().BeSameAs(enumerated);
    }

    [Fact]
    public void CheckBoxesは存在しないTagを指定した場合に失敗します()
    {
        using var document = Document.Open(EmptyDocumentPath, validate: true);

        var action = () => _ = document.CheckBoxes["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void CheckBoxesは同じTagのチェックボックスが複数存在する場合に失敗します()
    {
        using var document = Document.Open(DuplicateCheckBoxesPath, validate: true);

        var action = () => _ = document.CheckBoxes["Agreement"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void CheckBoxesはTagのないチェックボックスを列挙しません()
    {
        using var document = Document.Open(CheckBoxesWithoutTagPath, validate: true);

        document.CheckBoxes.Single().Tag.Should().Be("Agreement");
    }
}
