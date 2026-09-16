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

    [Fact]
    public void SelectedDateTimeプロパティは日付書式の二重引用符を表示し内側の年も変換します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "年を二重引用符で囲んだ日付選択Content Control.docx"));
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
            // w:dateFormatの値は"YYYY" yyyy、w:lidはen-US、w:calendarはgregorianです。
            // XMLの属性内では二重引用符を&quot;で記録しますが、書式としての文字は二重引用符です。
            // Wordで日付を選び直し、fullDateが2026-09-16T00:00:00Z、w:tが"2026" 2026となることを確認しました。
            // 単一引用符とは異なり、二重引用符は表示に残り、内側のYYYYも年に変わります。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。年を変えて保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("\"2027\" 2027");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは五つの月書式記号で月名と月番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "月書式記号を五つ並べた日付選択Content Control.docx"));
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
            // w:dateFormatはMMMMM、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月16日を選び直し、fullDateの更新とw:tのSeptember9を確認しました。
            // 五つのMを一つの月名指定とはせず、MMMMの月名とMの月番号を続けて表示します。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。12月へ変更した後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("December12");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは六つの月書式記号で月名と二桁の月番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "月書式記号を六つ並べた日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはMMMMMM、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月16日を選び直し、fullDateの更新とw:tのSeptember09を確認しました。
            // MMMMの月名にMMの二桁の月番号が続きます。MMMMMの月番号は一桁指定なので異なります。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。2月へ変更し、ゼロ埋めを含む表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("February02");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは五つの日書式記号で曜日名と日番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日書式記号を五つ並べた日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはddddd、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直し、fullDateが2026-09-03T00:00:00Z、w:tがThursday3となることを確認しました。
            // ddddの曜日名にdの日番号が続き、一桁の日をゼロ埋めしません。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。別の曜日の日付を設定して保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("Wednesday3");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは五つの大文字の日書式記号で曜日名と日番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "大文字の日書式記号を五つ並べた日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはDDDDD、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直し、fullDateが2026-09-03T00:00:00Z、w:tがThursday3となることを確認しました。
            // 大文字でもDDDDの曜日名にDの日番号が続き、一桁の日をゼロ埋めしません。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。別の曜日の日付を設定して保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("Wednesday3");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは六つの日書式記号で曜日名と二桁の日番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日書式記号を六つ並べた日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはdddddd、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直し、fullDateが2026-09-03T00:00:00Z、w:tがThursday03となることを確認しました。
            // ddddの曜日名にddの二桁の日番号が続きます。dddddの場合と異なり、一桁の日をゼロ埋めします。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。別の曜日の日付を設定して保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("Wednesday03");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは六つの大文字の日書式記号で曜日名と二桁の日番号を続けて表示します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "大文字の日書式記号を六つ並べた日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはDDDDDD、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直し、fullDateが2026-09-03T00:00:00Z、w:tがThursday03となることを確認しました。
            // 大文字でもDDDDの曜日名にDDの二桁の日番号が続き、一桁の日をゼロ埋めします。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。別の曜日の日付を設定して保存後の表示を確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("Wednesday03");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式のパーセントを表示に残します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日付書式にパーセントを含む日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはyyyy%MM%dd、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直し、fullDateが2026-09-03T00:00:00Z、w:tが2026%09%03となることを確認しました。
            // %は表示する文字であり、.NETの一文字用カスタム書式指定として扱いません。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。年月を変え、%と月日のゼロ埋めが残ることを確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("2027%02%03");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SelectedDateTimeプロパティは日付書式の小文字fを表示に残します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(
            Path.Combine("TestData", "日付書式に小文字fを含む日付選択Content Control.docx"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath, validate: true))
            {
                document.DatePickers.Single().SelectedDateTime =
                    new DateTimeOffset(2027, 2, 3, 0, 0, 0, TimeSpan.FromHours(9));
                document.SaveAs(outputPath);
            }

            using var saved = WordprocessingDocument.Open(outputPath, false);
            // w:dateFormatはyyyy-MM-dd f、w:lidはen-US、w:calendarはgregorianです。
            // Wordで9月3日を選び直すと、fullDateは2026-09-03T00:00:00Z、w:tは2026-09-03 fになります。
            // 小文字fは文字として残し、.NETの秒の小数部を表示する書式として解釈しません。
            // 固定DOCXは同じ書式を持つ最小OOXMLです。年月を変え、表示が更新されてもfが残ることを確認します。
            saved.MainDocumentPart!.Document!
                .Descendants<Wordprocessing.Text>()
                .Single().Text.Should().Be("2027-02-03 f");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    // 以下19件は入力を限定した未レビューの調査メモです。実装済みケースは含めません。
    // 特記がなければw:lid=en-US、w:calendar=gregorianを使います。
    // w:dateFormatは書式文字列、fullDateは保存日時、w:tは表示文字列です。
    // Wordで日付を選び直し、書式が保持され、fullDateとw:tが更新されたかを確認します。
    // 表示の期待値をレビュー後、固定DOCXと公開APIによる保存結果の検証に置き換えて解除します。
    // Wordが更新しない場合は無変更を仕様化せず、対応対象にするか確認してからメモを完了します。
    // 各ケースは有効化または対象外の記録で完了とし、別の未確認入力へ差し替えて残しません。

    [Fact(Skip = "yyyyyのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは五つの小文字年記号の表示を確認します()
    {
        // w:dateFormat=yyyyy、設定値=2027-02-03。五桁化と複数記号への分割の違いを確認します。
        // yyyは既に対象外と決定済みであり、その判断を再検討するテストではありません。
        throw new NotImplementedException("yyyyyのWord表示と対応方針を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "YYYYYのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは五つの大文字年記号の表示を確認します()
    {
        // w:dateFormat=YYYYY、設定値=2027-02-03。小文字版と同じ結果になるとは仮定しません。
        throw new NotImplementedException("YYYYYのWord表示と対応方針を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "hhhのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは三つの十二時間制の時記号の表示を確認します()
    {
        // w:dateFormat=hhh、設定値=2027-02-03T17:04:05+09:00。hhとhに分割されるかを確認します。
        // カレンダー操作だけでは時刻を指定できないため、Wordで時刻を反映できる確認手順も必要です。
        throw new NotImplementedException("hhhに対するWordの時刻表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "HHHのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは三つの二十四時間制の時記号の表示を確認します()
    {
        // w:dateFormat=HHH、設定値=2027-02-03T17:04:05+09:00。HHとHに分割されるかを確認します。
        throw new NotImplementedException("HHHに対するWordの時刻表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "mmmのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは三つの分記号の表示を確認します()
    {
        // w:dateFormat=mmm、設定値=2027-02-03T17:04:05+09:00。mmとmに分割されるかを確認します。
        throw new NotImplementedException("mmmに対するWordの時刻表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "sssのWord表示が未確認。受理と表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは三つの秒記号の表示を確認します()
    {
        // w:dateFormat=sss、設定値=2027-02-03T17:04:05+09:00。ssとsに分割されるかを確認します。
        throw new NotImplementedException("sssに対するWordの時刻表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "大文字FのWordでの意味が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の大文字Fの扱いを確認します()
    {
        // w:dateFormat=yyyy-MM-dd F、設定値=2027-02-03T00:00:00+09:00。
        // .NETではゼロの小数秒が表示されないため、Wordで文字として残るかを区別します。
        throw new NotImplementedException("大文字FのWord表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "大文字KのWordでの意味が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の大文字Kの扱いを確認します()
    {
        // w:dateFormat=yyyy-MM-dd K、設定値=2027-02-03T00:00:00+09:00。
        // .NETの時差表示へ置き換わることと、Wordの表示が一致するかを確認します。
        throw new NotImplementedException("大文字KのWord表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "zzzのWordでの意味が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の三つの小文字zの扱いを確認します()
    {
        // w:dateFormat=yyyy-MM-dd zzz、設定値=2027-02-03T00:00:00+09:00。
        // 時差の時分を表す.NETのzzzを代表例とします。zの長さを総当たりするメモではありません。
        throw new NotImplementedException("zzzのWord表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "小文字tのWordでの意味が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の小文字tの扱いを確認します()
    {
        // w:dateFormat=yyyy-MM-dd t、設定値=2027-02-03T00:00:00+09:00。
        // .NETでは午前午後名の先頭文字です。Wordでもそうなるとは仮定しません。
        throw new NotImplementedException("小文字tのWord表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "ttのWordでの意味が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式の二つの小文字tの扱いを確認します()
    {
        // w:dateFormat=yyyy-MM-dd tt、設定値=2027-02-03T00:00:00+09:00。
        // 文書が直接持つttを調べます。am/pmから変換した内部のttを調べるテストではありません。
        throw new NotImplementedException("文書が持つttのWord表示を確認してから検証コードを用意する。");
    }

    [Fact(Skip = "表示形式省略時の日本語表示が未確認。標準形式の根拠と期待値をレビューし、固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示形式を省略した日本語の日付表示を確認します()
    {
        // w:dateFormat要素を省略し、w:lid=ja-JP、w:calendar=gregorian、設定値=2027-02-03とします。
        // w:valの欠落や空文字とは区別します。実行環境のCurrentCultureを既定値と仮定しません。
        throw new NotImplementedException("ja-JPで表示形式を省略したときの標準形式を確認する。");
    }

    [Fact(Skip = "表示形式省略時の英語表示が未確認。標準形式の根拠と期待値をレビューし、固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示形式を省略した英語の日付表示を確認します()
    {
        // w:dateFormat要素を省略し、w:lid=en-US、w:calendar=gregorian、設定値=2027-02-03とします。
        // 日本語の既定書式へ固定していないことを確認する、表示言語の対照例です。
        throw new NotImplementedException("en-USで表示形式を省略したときの標準形式を確認する。");
    }

    [Fact(Skip = "表示言語省略時のrun言語参照が未確認。参照先と期待値をレビューし、固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語を省略してrunに英語を指定した場合の月名を確認します()
    {
        // w:lidを省略し、w:dateFormat=MMMM、内容runのw:rPr/w:langのw:val=en-US、設定値=2027-02-03。
        // 文書既定のrun言語はja-JPとし、直接指定を優先するかが表示から区別できるようにします。
        throw new NotImplementedException("w:lid省略時のrunの直接言語指定の扱いを確認する。");
    }

    [Fact(Skip = "表示言語省略時のスタイル言語参照が未確認。参照先と期待値をレビューし、固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語を省略して文字スタイルに英語を指定した場合の月名を確認します()
    {
        // w:lidとrunの直接w:langを省略し、w:dateFormat=MMMM、設定値=2027-02-03。
        // runのw:rStyleが参照する文字スタイルのw:rPr/w:langをen-US、文書既定はja-JPとします。
        // OOXMLの言語継承全般ではなく、この一段の参照だけを対象とします。
        throw new NotImplementedException("w:lid省略時の文字スタイルの言語指定の扱いを確認する。");
    }

    [Fact(Skip = "表示言語省略時の文書既定言語参照が未確認。参照先と期待値をレビューし、固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語を省略して文書既定に英語を指定した場合の月名を確認します()
    {
        // w:lid、runとスタイルのw:langを省略し、w:dateFormat=MMMM、設定値=2027-02-03。
        // styles.xmlのw:docDefaults/w:rPrDefault/w:rPr/w:langをen-USにします。
        // プロセスのカルチャーをja-JPにしても文書の指定に従うかを確認し、変更したカルチャーは復元します。
        throw new NotImplementedException("w:lid省略時の文書既定の言語指定の扱いを確認する。");
    }

    [Fact(Skip = "暦省略時のWord表示が未確認。表示を確認し、期待値レビューと固定DOCX準備後に解除する。")]
    public void SelectedDateTimeプロパティは暦を省略した場合の年月日表示を確認します()
    {
        // w:calendarだけを省略し、w:dateFormat=yyyy/MM/dd、w:lid=ja-JP、設定値=2027-02-03。
        // 明示的なgregorianの既存テストと対照し、省略が表示へ影響するかを確認します。
        throw new NotImplementedException("w:calendar省略時のWord表示を確認する。");
    }

    [Fact(Skip = "japan暦指定時のWord表示が未確認。UIの暦と保存表示を分けて確認し、期待値レビュー後に解除する。")]
    public void SelectedDateTimeプロパティは日本の暦を指定した場合の年月日表示を確認します()
    {
        // w:calendar=japan、w:dateFormat=yyyy/MM/dd、w:lid=ja-JP、設定値=2027-02-03。
        // カレンダーUIが和暦でも、保存するw:tの年まで和暦になるとは仮定しません。
        // 元号記号gや和暦年記号eの対応は別の範囲検討課題で、このメモへ追加しません。
        throw new NotImplementedException("japan暦指定と保存する表示文字列の関係を確認する。");
    }

    [Fact(Skip = "空の表示形式の扱いが未決定。Wordの受理と更新を調査し、ライブラリの契約レビュー後に解除する。")]
    public void SelectedDateTimeプロパティは表示形式が空文字の場合の扱いを確認します()
    {
        // w:dateFormat要素とw:val属性は存在し、値だけ空文字。w:lid=en-US、設定値=2027-02-03。
        // 要素省略とは別の入力です。Wordの結果を確認する前に既定書式への変換や例外を仕様化しません。
        throw new NotImplementedException("空の表示形式の受理とライブラリでの扱いを確認する。");
    }

}
