using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class DatePickerCollectionのテスト
{
    static readonly string EmptyDocumentPath =
        Path.Combine("TestData", "空の文書.docx");

    static readonly string DatePickerContentControlPath =
        Path.Combine("TestData", "日付選択のContent Control.docx");

    static readonly string DatePickersWithoutTagPath =
        Path.Combine("TestData", "Tagなしを含む日付選択Content Control.docx");

    static readonly string DuplicateDatePickersPath =
        Path.Combine("TestData", "同じTagの日付選択Content Control.docx");

    [Fact]
    public void DatePickersは文書内の日付選択ContentControlを列挙します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers.Should().ContainSingle();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void DatePickersはTagから日付選択ContentControlを取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers["DeliveryDate"].Tag.Should().Be("DeliveryDate");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void DatePickersは列挙とTag検索で同じ日付選択ContentControlを返します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var enumerated = document.DatePickers.Single();

            document.DatePickers["DeliveryDate"].Should().BeSameAs(enumerated);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void DatePickersは存在しないTagを指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () => _ = document.DatePickers["not_found"];

            action.Should().Throw<KeyNotFoundException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void DatePickersは同じTagの日付選択ContentControlが複数存在する場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DuplicateDatePickersPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () => _ = document.DatePickers["DeliveryDate"];

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void DatePickersはTagのない日付選択ContentControlを列挙しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(DatePickersWithoutTagPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers.Single().Tag.Should().Be("DeliveryDate");
        }
        finally
        {
            File.Delete(filePath);
        }
    }
}
