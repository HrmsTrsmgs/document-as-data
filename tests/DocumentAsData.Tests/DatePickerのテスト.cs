using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test;

public class DatePickerのテスト
{
    static readonly string DatePickerContentControlPath =
        Path.Combine("TestData", "日付選択のContent Control.docx");

    static readonly string OffsetDatePickerContentControlPath =
        Path.Combine("TestData", "時差付き日付選択Content Control.docx");

    [Fact]
    public void Documentプロパティは日付選択ContentControlが属する文書を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers.Single().Document.Should().BeSameAs(document);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Tagプロパティは日付選択ContentControlのTagを取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
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

    [Fact]
    public void ToStringは種類と引用符で囲んだTagを返します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers.Single().ToString().Should().Be(
                "DatePicker { Tag = \"DeliveryDate\" }");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付選択ContentControlの日時を取得します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.DatePickers.Single().SelectedDateTime
                .Should().Be(
                    new DateTimeOffset(
                        2026,
                        9,
                        4,
                        0,
                        0,
                        0,
                        TimeSpan.Zero));
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティはfullDateの時差を保持します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(OffsetDatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var expected = new DateTimeOffset(
                2026,
                9,
                4,
                0,
                0,
                0,
                TimeSpan.FromHours(9));

            // OOXMLではfullDateに完全なXML Schema DateTimeを保持します。
            // 2026-09-04T00:00:00+09:00の日時と時差をそのまま読み取ります。
            var actual = document.DatePickers.Single().SelectedDateTime;

            actual.Should().Be(expected);
            actual.Offset.Should().Be(expected.Offset);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付選択ContentControlの日時を設定します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);
            var tested = document.DatePickers.Single();
            var value = new DateTimeOffset(
                2026,
                12,
                31,
                0,
                0,
                0,
                TimeSpan.FromHours(9));

            tested.SelectedDateTime = value;

            tested.SelectedDateTime.Should().Be(value);
            tested.SelectedDateTime.Offset.Should().Be(value.Offset);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは表示形式に従った日付文字列を設定します()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(
                        2026,
                        12,
                        31,
                        0,
                        0,
                        0,
                        TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // 日付の値はw:date/@w:fullDate、Wordで表示される文字列はw:tに別々に保持されます。
            // テンプレートのw:dateFormat="yyyy/MM/dd"に従ってw:tも更新します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("2026/12/31");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }
}
