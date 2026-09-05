using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class MergeFieldCollectionのテスト
{
    static readonly string TestFilePath =
        Path.Combine("TestData", "単純形式のMERGEFIELD.docx");
    static readonly string DuplicateMergeFieldsPath =
        Path.Combine("TestData", "同名のMERGEFIELD.docx");
    static readonly string StandardComplexMergeFieldPath =
        Path.Combine(
            "TestData",
            "Word標準の書式維持スイッチを持つ複合MERGEFIELD.docx");
    static readonly string UnquotedComplexMergeFieldPath =
        Path.Combine(
            "TestData",
            "引用符なしの書式維持スイッチを持つ複合MERGEFIELD.docx");
    static readonly string MultipleComplexMergeFieldsPath =
        Path.Combine("TestData", "複数の複合MERGEFIELD.docx");
    static readonly string DateFieldPath =
        Path.Combine("TestData", "DATEフィールド.docx");

    [Fact]
    public void MergeFieldsは文書内のMERGEFIELDを列挙します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            document.MergeFields
                .Select(it => it.Name)
                .Should().Equal("CustomerName", "Address");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは名前からMERGEFIELDを取得します()
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
    public void MergeFieldsは列挙と名前検索で同じMERGEFIELDを返します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);
            var enumerated = (
                from field in document.MergeFields
                where field.Name == "CustomerName"
                select field
            ).Single();

            document.MergeFields["CustomerName"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは存在しない名前を指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(TestFilePath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.MergeFields["not_found"];

            action.Should().Throw<KeyNotFoundException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは同名のMERGEFIELDが複数存在する場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () => _ = document.MergeFields["CustomerName"];

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsはMERGEFORMATスイッチをMERGEFIELD名に含めません()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(StandardComplexMergeFieldPath);
        try
        {
            using var document = Document.Open(filePath);

            // 複合フィールドは、w:fldCharのbegin、w:instrText、separate、
            // 表示結果、endという要素列で保存されます。この固定データの
            // w:instrTextは次のWord標準の命令を保持しています。
            //   MERGEFIELD "CustomerName" \* MERGEFORMAT
            // MERGEFORMATは差し込み後も結果の書式を維持するスイッチであり、
            // 引用符で囲まれたCustomerNameだけがMERGEFIELD名です。
            document.MergeFields.Single().Name.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは引用符なしの名前にMERGEFORMATスイッチを含めません()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(UnquotedComplexMergeFieldPath);
        try
        {
            using var document = Document.Open(filePath);

            // MERGEFIELD名は引用符なしでも保存できます。固定データの
            // w:instrTextは次の命令を保持しています。
            //   MERGEFIELD CustomerName \* MERGEFORMAT
            // 空白に続くMERGEFORMATはフィールドの書式維持スイッチであり、
            // CustomerNameだけがMERGEFIELD名です。
            document.MergeFields.Single().Name.Should().Be("CustomerName");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsは複数の複合MERGEFIELDを文書内の順序で列挙します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(MultipleComplexMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            // 固定データには、begin、instrText、separate、表示結果、endの
            // 要素列で構成された複合MERGEFIELDが二つ続けて保存されています。
            // 一つ目のw:instrTextはCustomerName、二つ目はAddressです。
            // 一要素で完結するw:fldSimpleとは異なり、最初のendまで読み終えた
            // 状態を次のbeginへ持ち越さず、二つを文書順に列挙することを
            // 確認します。
            document.MergeFields
                .Select(it => it.Name)
                .Should().Equal("CustomerName", "Address");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void MergeFieldsはDATEフィールドを列挙しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DateFieldPath);
        try
        {
            using var document = Document.Open(filePath);

            // 固定データには、MERGEFIELDと同じくbegin、instrText、
            // separate、表示結果、endの要素列で構成されたDATEフィールドが
            // 保存されています。w:instrTextの命令は次の内容です。
            //   DATE \@ "yyyy/MM/dd"
            // 複合フィールドという内部構造が同じでも、命令の種類がDATEなら
            // 名前付きデータ項目であるMERGEFIELDとして扱わないことを
            // 確認します。
            document.MergeFields.Should().BeEmpty();
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
