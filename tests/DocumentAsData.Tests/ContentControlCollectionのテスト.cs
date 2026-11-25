using FluentAssertions;

namespace Marimo.DocumentAsData.Test;

public class ContentControlCollectionのテスト
{
    static readonly string EmptyDocumentPath =
        Path.Combine("TestData", "空の文書.docx");
    static readonly string ContentControlPath =
        Path.Combine("TestData", "単一のContent Control.docx");
    static readonly string ContentControlsPath =
        Path.Combine("TestData", "複数のContent Control.docx");
    static readonly string ContentControlsWithoutTagPath =
        Path.Combine("TestData", "Tagなしを含むContent Control.docx");
    static readonly string NestedContentControlsPath =
        Path.Combine("TestData", "ネストしたContent Control.docx");
    static readonly string DuplicateContentControlsPath =
        Path.Combine("TestData", "同じTagのContent Control.docx");
    static readonly string CheckBoxContentControlPath =
        Path.Combine("TestData", "チェックボックスのContent Control.docx");
    static readonly string DateContentControlPath =
        Path.Combine("TestData", "日付選択のContent Control.docx");
    static readonly string PictureContentControlPath =
        Path.Combine("TestData", "画像のContent Control.docx");
    static readonly string DropDownListContentControlPath =
        Path.Combine("TestData", "ドロップダウンリストのContent Control.docx");
    static readonly string ComboBoxContentControlPath =
        Path.Combine("TestData", "コンボボックスのContent Control.docx");
    static readonly string RepeatingSectionContentControlPath =
        Path.Combine("TestData", "繰り返しセクションのContent Control.docx");
    static readonly string EquationContentControlPath =
        Path.Combine("TestData", "数式のContent Control.docx");
    static readonly string CitationContentControlPath =
        Path.Combine("TestData", "引用のContent Control.docx");
    static readonly string GroupContentControlPath =
        Path.Combine("TestData", "グループのContent Control.docx");
    static readonly string BibliographyContentControlPath =
        Path.Combine("TestData", "文献目録のContent Control.docx");
    static readonly string DocumentPartContentControlPath =
        Path.Combine("TestData", "組み込み文書パーツのContent Control.docx");
    static readonly string DocumentPartListContentControlPath =
        Path.Combine("TestData", "文書パーツギャラリーのContent Control.docx");
    static readonly string EntityPickerContentControlPath =
        Path.Combine("TestData", "エンティティピッカーのContent Control.docx");

    [Fact]
    public void ContentControlsは文書内のContentControlを列挙します()
    {
        using var document = Document.Open(ContentControlsPath);

        document.ContentControls.Should().HaveCount(2);
    }

    [Fact]
    public void ContentControlsは文書内の順序で列挙します()
    {
        using var document = Document.Open(ContentControlsPath);

        document.ContentControls
            .Select(it => it.Tag)
            .Should().Equal("CustomerName", "Address");
    }

    [Fact]
    public void ContentControlsはTagからContentControlを取得します()
    {
        using var document = Document.Open(ContentControlPath);

        document.ContentControls["CustomerName"].Tag.Should().Be("CustomerName");
    }

    [Fact]
    public void ContentControlsは列挙とTag検索で同じContentControlを返します()
    {
        using var document = Document.Open(ContentControlPath);
        var enumerated = document.ContentControls.Single();

        document.ContentControls["CustomerName"].Should().BeSameAs(enumerated);
    }

    [Fact]
    public void ContentControlsは存在しないTagを指定した場合に失敗します()
    {
        using var document = Document.Open(EmptyDocumentPath);

        var action = () => _ = document.ContentControls["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void ContentControlsは同じTagのContentControlが複数存在する場合に失敗します()
    {
        using var document = Document.Open(DuplicateContentControlsPath);

        var action = () => _ = document.ContentControls["CustomerName"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void ContentControlsはTagのないContentControlを列挙しません()
    {
        using var document = Document.Open(ContentControlsWithoutTagPath);

        document.ContentControls.Single().Tag.Should().Be("CustomerName");
    }

    [Fact]
    public void ContentControlsはネストしたContentControlをそれぞれ列挙します()
    {
        using var document = Document.Open(NestedContentControlsPath);

        document.ContentControls
            .Select(it => it.Tag)
            .Should().Equal("Outer", "Inner");
    }

    [Fact]
    public void ContentControlsはチェックボックスを列挙しません()
    {
        using var document = Document.Open(CheckBoxContentControlPath, validate: true);

        // Wordのチェックボックスは、表示記号のw:tとは別に
        // w:sdtPr/w14:checkbox/w14:checkedへチェック状態を保持します。
        // 文字列用のTextプロパティでは両者を同期できないため、通常のContentControlsには含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは日付選択ContentControlを列挙しません()
    {
        using var document = Document.Open(DateContentControlPath, validate: true);

        // Wordの日付選択は、w:sdtContent/w:tの表示文字列とは別に
        // w:sdtPr/w:dateへ日付、表示形式、言語、暦を保持します。
        // 文字列用のTextプロパティでは両者を同期できないため、通常のContentControlsには含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは画像ContentControlを列挙しません()
    {
        using var document = Document.Open(PictureContentControlPath, validate: true);

        // Wordの画像Content Controlはw:sdtPr/w:pictureで種類を示し、
        // w:sdtContent内のw:drawingと関連する画像Partへ画像を保持します。
        // 文字列用のTextプロパティでは扱えないため、通常のContentControlsには含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsはドロップダウンリストを列挙しません()
    {
        using var document = Document.Open(DropDownListContentControlPath, validate: true);

        // Wordのドロップダウンリストは、w:sdtContent/w:tの表示文字列とは別に
        // w:sdtPr/w:dropDownListへ選択肢と最後に選択された値を保持します。
        // 文字列用のTextプロパティでは両者を同期できないため、通常のContentControlsには含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsはコンボボックスを列挙しません()
    {
        using var document = Document.Open(ComboBoxContentControlPath, validate: true);

        // Wordのコンボボックスは、w:sdtContent/w:tの表示文字列とは別に
        // w:sdtPr/w:comboBoxへ選択肢と最後に選択された値を保持します。
        // 文字列用のTextプロパティでは両者を同期できないため、通常のContentControlsには含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは繰り返しセクションのコンテナを列挙しません()
    {
        using var document = Document.Open(RepeatingSectionContentControlPath, validate: true);

        // w15:repeatingSectionは繰り返し全体、w15:repeatingSectionItemは
        // 1件分の領域を束ねるContent Controlで、どちらも値ではありません。
        // それらの内側にある通常の入力項目だけを列挙します。
        document.ContentControls
            .Select(it => it.Tag)
            .Should().Equal("ItemName");
    }

    [Fact]
    public void ContentControlsは繰り返し内に1件だけあるTagを文書全体から取得できます()
    {
        using var document = Document.Open(RepeatingSectionContentControlPath);

        document.ContentControls["ItemName"].Text.Should().Be("商品A");
    }

    [Fact]
    public void ContentControlsは繰り返し内の複数の明細に同じTagがあると失敗します()
    {
        using var document = Document.Open(
            Path.Combine("TestData", "繰り返しセクションに2件の明細.docx"));

        var action = () => document.ContentControls["ItemName"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void ContentControlsは数式を列挙しません()
    {
        using var document = Document.Open(EquationContentControlPath, validate: true);

        // w:sdtPr/w:equationは数式Content Controlを示し、内容は
        // w:tではなくOffice Mathのm:oMath/m:r/m:tとして保持されます。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは引用を列挙しません()
    {
        using var document = Document.Open(CitationContentControlPath, validate: true);

        // w:sdtPr/w:citationはWordが引用として管理するContent Controlを示します。
        // 表示文字列のみを値とする通常項目には含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsはグループのコンテナを列挙しません()
    {
        using var document = Document.Open(GroupContentControlPath, validate: true);

        // w:sdtPr/w:groupは複数の内容を束ねるコンテナで、値ではありません。
        // 内側にある通常の入力項目だけを列挙します。
        document.ContentControls
            .Select(it => it.Tag)
            .Should().Equal("CustomerName");
    }

    [Fact]
    public void ContentControlsは文献目録を列挙しません()
    {
        using var document = Document.Open(BibliographyContentControlPath, validate: true);

        // w:sdtPr/w:bibliographyはWordが文献目録として管理する領域を示します。
        // 表示文字列のみを値とする通常項目には含めません。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは組み込み文書パーツを列挙しません()
    {
        using var document = Document.Open(DocumentPartContentControlPath, validate: true);

        // w:sdtPr/w:docPartObjは組み込み文書パーツを表し、
        // w:docPartGalleryなどの文書パーツ用情報を保持します。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsは文書パーツギャラリーを列挙しません()
    {
        using var document = Document.Open(DocumentPartListContentControlPath, validate: true);

        // w:sdtPr/w:docPartListは文書パーツの選択肢を表し、
        // w:docPartGalleryなどの文書パーツ用情報を保持します。
        document.ContentControls.Should().BeEmpty();
    }

    [Fact]
    public void ContentControlsはエンティティピッカーを列挙しません()
    {
        using var document = Document.Open(EntityPickerContentControlPath, validate: true);

        // w:sdtPr/w14:entityPickerはエンティティ選択用のContent Controlを示します。
        // 選択状態を表示文字列だけで表せないため、通常項目には含めません。
        document.ContentControls.Should().BeEmpty();
    }
}
