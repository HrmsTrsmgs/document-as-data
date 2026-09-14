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
    public void SelectedDateTimeプロパティは未入力のコントロールに日時を設定できます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日時が未入力の日付選択Content Control.docx"));
        try
        {
            // w:dateはありますが、日時を保持するw:fullDate属性はなく、
            // 表示文字列のw:tも空です。書式と表示言語は設定されています。
            // 未入力状態を狙った最小OOXMLで、Wordのプレースホルダー表示は対象外です。
            using var document = Document.Open(filePath, validate: true);
            var tested = document.DatePickers.Single();
            var value = new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));

            tested.SelectedDateTime = value;

            tested.SelectedDateTime.Should().Be(value);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティはMMYYYY形式の年を四桁で表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "表示形式がMM-YYYYの日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:date/@w:fullDateは日時そのもの、w:dateFormatはWord用の表示形式、
            // w:tは実際に文書に表示する文字列です。
            // OOXMLのMM-YYYYは月と四桁の年を表しますが、.NETの書式へそのまま
            // 渡すとYYYYが文字のまま残るため、保存した表示文字列を確認します。
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.dateformat
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("12-2026");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは引用された文字列を残して大文字の年指定を表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "表示形式に大文字の年と引用文字列を含む日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatの「'YYYY' YYYY/MM/dd」では、引用符の内側は文字列、
            // 外側のYYYYは四桁の年です。w:tに保存した表示内容を検証します。
            // .NETにそのまま渡しても、Yを一括で小文字にしても正しく表示できません。
            // 固定DOCXはこの書式を指定したOOXMLであり、Word実機での確認資料ではありません。
            // https://support.microsoft.com/en-us/word/format-field-results
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("YYYY 2026/12/31");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
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

    [Theory]
    [InlineData("表示形式が小文字の二桁年の日付選択Content Control.docx")]
    [InlineData("表示形式が大文字の二桁年の日付選択Content Control.docx")]
    public void SelectedDateTimeプロパティは二桁の年を表示します(string fileName)
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(Path.Combine("TestData", fileName));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2006, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatは小文字のyyまたは大文字のYY、w:lidはja-JP、
            // w:calendarはgregorianです。2006年を二桁の06で表示することを、
            // 保存されたw:sdtContent内のw:t（表示文字列）で確認します。
            // 固定DOCXは書式を指定した最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("06");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは大文字の日指定を日として表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "表示形式がYYYY MM DDの日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはYYYY/MM/DD、w:lidはja-JP、w:calendarはgregorianです。
            // WordのDDは二桁の日ですが、.NETにそのまま渡すと文字列DDが残ります。
            // w:sdtContent内のw:t（保存された表示文字列）で日への変換を確認します。
            // 固定DOCXは書式を指定した最小OOXMLで、Word実機による確認資料ではありません。
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

    [Theory]
    [InlineData("表示形式がDDDで英語の曜日を表示する日付選択Content Control.docx", "Thu")]
    [InlineData("表示形式がDDDDで英語の曜日を表示する日付選択Content Control.docx", "Thursday")]
    public void SelectedDateTimeプロパティは大文字の曜日指定を曜日として表示します(string fileName, string expectedText)
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(Path.Combine("TestData", fileName));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはDDD（曜日の略称）またはDDDD（曜日の正式名称）、
            // w:lidはen-US、w:calendarはgregorianです。日を数値で表示するDDと異なり、
            // 木曜日をThuまたはThursdayとしてw:sdtContent内のw:tへ保存することを確認します。
            // 固定DOCXは書式を指定した最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(expectedText);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは月と日の桁数指定に従ってゼロ埋めします()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "月日を一桁指定と二桁指定で表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatは「M/d MM/dd」、w:lidはja-JP、w:calendarはgregorianです。
            // Mとdはゼロ埋めせず、MMとddは二桁にします。どちらも一桁になる2月3日を使い、
            // w:sdtContent内のw:t（保存された表示文字列）で桁数指定の違いを確認します。
            // 固定DOCXは書式を指定した最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("2/3 02/03");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは表示言語に従った月名を表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "英語の月名を略称と正式名称で表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはMMM MMMM、w:lidはen-US、w:calendarはgregorianです。
            // MMMは月名の略称、MMMMは正式名称で、2月はFeb Februaryとなります。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("Feb February");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは表示言語に従った曜日名を表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日本語の曜日を略称と正式名称で表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはddd dddd、w:lidはja-JP、w:calendarはgregorianです。
            // dddは曜日の略称、ddddは正式名称で、日本語の木曜日は木 木曜日となります。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("木 木曜日");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData(5, "5 05 5 05")]
    [InlineData(17, "5 05 17 17")]
    public void SelectedDateTimeプロパティは時刻を十二時間制と二十四時間制で表示します(int hour, string expectedText)
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "時刻を十二時間制と二十四時間制で表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, hour, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはh hh H HH、w:lidはja-JP、w:calendarはgregorianです。
            // hは十二時間制、Hは二十四時間制で、二文字指定は二桁にゼロ埋めします。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(expectedText);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは分と秒の桁数指定に従って表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "月と分を区別して分秒の桁数を指定する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, 17, 4, 5, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはM m mm s ss、w:lidはja-JP、w:calendarはgregorianです。
            // Mは月、mは分、sは秒です。mmとssだけ二桁へゼロ埋めします。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("2 4 04 5 05");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData(5, "午前 午前")]
    [InlineData(17, "午後 午後")]
    public void SelectedDateTimeプロパティは表示言語に従って午前と午後を表示します(int hour, string expectedText)
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日本語の午前午後を大小文字の指定で表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, hour, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはam/pm AM/PM、w:lidはja-JP、w:calendarはgregorianです。
            // am/pmとAM/PMはいずれも午前午後の指定です。表示言語に従って午前・午後を表示し、
            // 個々のmを分、Mを月、スラッシュを区切り文字として扱いません。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(expectedText);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData("日だけを一文字の書式で表示する日付選択Content Control.docx", "3")]
    [InlineData("月だけを一文字の書式で表示する日付選択Content Control.docx", "2")]
    [InlineData("時だけを一文字の書式で表示する日付選択Content Control.docx", "5")]
    public void SelectedDateTimeプロパティは一文字の書式をその項目だけの表示として扱います(string fileName, string expectedText)
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(Path.Combine("TestData", fileName));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 2, 3, 17, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormat全体がd（日）、M（月）、h（十二時間制の時）の一文字です。
            // w:lidはja-JP、w:calendarはgregorianです。.NETの標準書式には切り替えず、
            // 指定された一項目だけをw:sdtContent内のw:t（表示文字列）へ保存します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(expectedText);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式のスラッシュを地域別の区切り文字へ変更しません()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "ドイツ語設定で日付をスラッシュ区切りにする日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはyyyy/MM/dd、w:lidはde-DE、w:calendarはgregorianです。
            // .NETでは/が地域別の日付区切り文字になりますが、ここでは書式の/を保持します。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
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

    [Fact]
    public void SelectedDateTimeプロパティは時刻書式のコロンを地域別の区切り文字へ変更しません()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "フィンランド語設定で時刻をコロン区切りにする日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 17, 4, 5, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはHH:mm:ss、w:lidはfi-FI、w:calendarはgregorianです。
            // .NETでは:が地域別の時刻区切り文字になりますが、ここでは書式の:を保持します。
            // w:sdtContent内のw:t（保存された表示文字列）を検証します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("17:04:05");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式内の閉じられていない単一引用符を表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "二桁年の前に単一引用符を表示する日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2026, 12, 31, 17, 4, 5, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはHH:mm MMM-d, 'yy、w:lidはen-USです。
            // Wordの公式例では、対にならない'は表示文字で、その後のyyは二桁年です。
            // https://support.microsoft.com/en-us/word/format-field-results
            // .NETの閉じられていない引用文字列として例外にせず、保存後のw:tで表示を確認します。
            // 固定DOCXは最小OOXMLで、Word実機による確認資料ではありません。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("17:04 Dec-31, '26");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式の単一引用符内のバックスラッシュを表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "引用文字列にバックスラッシュを含む日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatは'C:\Temp' yyyy、w:lidはen-US、w:calendarはgregorianです。
            // 単一引用符は表示文字列を囲み、外側のyyyyだけが年へ置き換わります。
            // Wordでカレンダーから日付を選び直し、保存後のw:tがC:\Temp 2026となることを確認済みです。
            // .NETは引用符内でもバックスラッシュをエスケープとして扱うため、そのまま渡すと消えてしまいます。
            // 固定DOCXは最小OOXMLで作成し、ここでは年も変えて既存の表示文字列のままでないことを確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(@"C:\Temp 2027");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式の引用符外のバックスラッシュを表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "年月日をバックスラッシュで区切る日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 12, 31, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはyyyy\MM\dd、w:lidはen-US、w:calendarはgregorianです。
            // Wordで日付を選び直し、fullDateが2026-09-16T00:00:00Z、w:tが2026\09\16となることを確認しました。
            // バックスラッシュは表示する区切り文字であり、直後のMやdを書式記号から除外しません。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。年・月・日を変えて保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be(@"2027\12\31");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    // 以下は未レビューの保留メモです。書式の全組み合わせではなく解釈ルールごとに置きます。
    // w:dateFormat（表示形式）、w:lid（言語）、w:calendar（暦）を持つ固定DOCXを使い、
    // 公開APIで日時を設定・保存した後のw:t（表示文字列）を確認する予定です。
    // 未確定のWordの挙動や例外契約は、この段階で断定しません。
    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式内の引用符とエスケープの扱いを確認します()
    {
        // 単一引用符内のYYYYと、閉じられていない単一引用符は確認済み。
        // 単一引用符内と引用符外のバックスラッシュは確認済み。二重引用符は未確認。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の記号の繰り返しの扱いを確認します()
    {
        // yyyなどの連続指定について、Wordが記号を区切る規則と.NETとの差を確認して期待値を決める。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の通常文字をNET固有の書式として解釈しません()
    {
        // Wordで通常文字となるもののうち、.NETでは特別な意味を持つ文字を選ぶ。対象文字はWordでの確認後に決める。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示形式が省略された場合の表示を確認します()
    {
        // w:dateFormatがない場合の言語に基づく表示形式を確認する。欠落した属性値や不正形式とは分ける。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語が省略された場合の表示を確認します()
    {
        // w:lidがない場合に参照する内容のrunの言語を確認する。実行環境の現在カルチャーで勝手に補わない。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは暦指定と表示言語による表示の扱いを確認します()
    {
        // w:calendarは日付選択UIの暦指定でもある。保存表示への影響をWordで確認し、UI設定から表示の年を推測しない。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは非対応の日付書式を指定された場合の扱いを確認します()
    {
        // 対応範囲と、非対応時に例外・無変更などのどの契約を採るかを先にレビューする。現段階で例外型を固定しない。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

}
