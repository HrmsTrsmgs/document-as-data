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

    // 以下は未レビューの保留メモです。書式の全組み合わせではなく解釈ルールごとに置きます。
    // w:dateFormat（表示形式）、w:lid（言語）、w:calendar（暦）を持つ固定DOCXを使い、
    // 公開APIで日時を設定・保存した後のw:t（表示文字列）を確認する予定です。
    // 未確定のWordの挙動や例外契約は、この段階で断定しません。
    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語に従った月名を表示します()
    {
        // MMMとMMMMの略称・正式名称を、月名が区別できる言語で確認する。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語に従った曜日名を表示します()
    {
        // dddとddddの略称・正式名称を確認する。大文字からの変換ではなく表示言語が対象。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは時刻を十二時間制と二十四時間制で表示します()
    {
        // 午後の時刻でhとHの違いを確認し、hhとHHのゼロ埋めも同じ時刻表示の観点として扱う。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは分と秒の桁数指定に従って表示します()
    {
        // 一桁の分秒でmとmm、sとssを確認する。月のMと分のmを混同しない。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは表示言語に従って午前と午後を表示します()
    {
        // am/pmとAM/PMを確認する。午前・午後それぞれの代表値を使い、.NETのttとの対応を検証する。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは一文字の書式をその項目だけの表示として扱います()
    {
        // 日だけのd、月だけのM、時だけのhなどを確認する。.NETの標準書式への切替を防ぐ観点。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式のスラッシュを地域別の区切り文字へ変更しません()
    {
        // 地域設定の区切り文字がスラッシュでない言語を選び、書式で指定した文字が残ることを確認する。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは時刻書式のコロンを地域別の区切り文字へ変更しません()
    {
        // 地域設定の区切り文字がコロンでない言語を選び、書式で指定した文字が残ることを確認する。
        throw new NotImplementedException("保留メモのため、固定DOCXと検証コードを準備してから有効化する。");
    }

    [Fact(Skip = "日付書式の保留メモ。0.3.0への採用は未決定。仕様レビューと固定DOCX・検証コードの準備後に解除する。")]
    public void SelectedDateTimeプロパティは日付書式内の引用符とエスケープの扱いを確認します()
    {
        // 単一引用符内のYYYYを保持する基本例は既存テストで確認済み。引用符そのもの、二重引用符、バックスラッシュはWordで意味を確認してから期待値を決める。
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
