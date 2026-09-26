using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の日付選択Content Controlを表します。
/// </summary>
public partial class DatePicker : ContentControl
{
    /// <summary>
    /// 日付選択Content Controlを初期化します。
    /// </summary>
    /// <param name="document">日付選択Content Controlが属する文書。</param>
    /// <param name="element">日付選択Content Controlを構成するOOXML要素。</param>
    internal DatePicker(
        Document document,
        Wordprocessing.SdtElement element)
        : base(document, element)
    {
    }

    /// <summary>
    /// 種類と、引用符で囲んだTagを返します。
    /// </summary>
    /// <returns>種類とTagを確認できる表示文字列。</returns>
    public override string ToString() =>
        $"DatePicker {{ Tag = {QuoteName(Tag)} }}";

    /// <summary>
    /// 日付選択Content Controlの日時を取得または設定します。
    /// </summary>
    /// <remarks>
    /// 日時が未入力でも設定できます。表示形式の省略・空文字には、表示言語の短い日付形式を使います。
    /// 表示言語の省略時は、表示runの直接指定、文字スタイル、文書既定の順に言語を参照します。
    /// 言語を決定できる設定と表示文字列は必要です。Wordの全書式に対応するものではありません。
    /// </remarks>
    /// <exception cref="InvalidOperationException">
    /// 取得時に日時が欠落している場合。または日付選択のプロパティや、設定時に必要な
    /// 表示設定の値または表示文字列が欠落しているか、必要な要素が複数存在する場合。
    /// </exception>
    /// <exception cref="NotSupportedException">
    /// このContent Controlまたは親の繰り返しセクションがWordのXMLマッピングを使用している場合。
    /// </exception>
    public DateTimeOffset SelectedDateTime
    {
        get
        {
            EnsureNotDataBound();
            return XmlConvert.ToDateTimeOffset(
                DateProperties.FullDate?.InnerText
                    ?? throw new InvalidOperationException());
        }
        set
        {
            EnsureNotDataBound();
            var properties = DateProperties;
            var culture = ResolveDisplayCulture(properties);
            var format = ResolveDisplayFormat(properties, culture);

            properties.FullDate = new() { InnerText = XmlConvert.ToString(value) };
            DisplayText.Text = value.ToString(
                ToDotNetDateFormat(format),
                culture);
        }
    }

    /// <summary>
    /// 日付選択自身、表示run、文字スタイル、文書既定の順に、最初に指定された表示言語を使います。
    /// 後順位のOOXML要素は、先順位で言語が決まらない場合だけ参照します。
    /// </summary>
    /// <param name="properties">更新対象の日付選択の設定。</param>
    /// <returns>表示言語に対応するカルチャー。</returns>
    CultureInfo ResolveDisplayCulture(Wordprocessing.SdtContentDate properties) =>
        CultureInfo.GetCultureInfo(
            properties.LanguageId?.Val?.Value
                ?? DisplayRunProperties?.Languages?.Val?.Value
                ?? DisplayCharacterStyle?.StyleRunProperties?.Languages?.Val?.Value
                ?? DocumentStyles?.DocDefaults?.RunPropertiesDefault
                    ?.RunPropertiesBaseStyle?.Languages?.Val?.Value
                ?? throw new InvalidOperationException());

    /// <summary>
    /// 表示形式の省略と空文字には表示言語の短い日付形式を使います。
    /// 表示形式要素があるのに値が欠落している場合は、不正な設定として扱います。
    /// </summary>
    /// <param name="properties">更新対象の日付選択の設定。</param>
    /// <param name="culture">既定の表示形式を決める表示言語。</param>
    /// <returns>表示に使用する日付書式。</returns>
    static string ResolveDisplayFormat(Wordprocessing.SdtContentDate properties, CultureInfo culture)
    {
        var format = properties.DateFormat is { } dateFormat
            ? dateFormat.Val?.Value ?? throw new InvalidOperationException()
            : culture.DateTimeFormat.ShortDatePattern;

        return format.IsEmpty()
            ? culture.DateTimeFormat.ShortDatePattern
            : format;
    }

    /// <summary>
    /// 単一引用符で囲まれた表示文字列を保持し、Wordの年・日・午前午後指定を.NET用に変換します。
    /// 一文字の書式も、.NETの標準書式ではなく指定された一項目として扱います。
    /// スラッシュとコロンは地域別の区切り文字へ変更せず、そのまま表示します。
    /// 対にならない単一引用符は文字として表示します。
    /// 二重引用符は文字として表示し、その内側も書式記号として扱います。
    /// バックスラッシュは.NET用にエスケープして表示に残します。
    /// パーセントは.NETの書式指定として解釈せず、表示に残します。
    /// f・Fも秒の小数部として解釈せず、表示に残します。
    /// K・zzz・t・ttも.NETの時差や午前午後の記号として解釈せず、表示に残します。
    /// yyyyy・YYYYYは四桁年と年の下二桁を続けて表示します。
    /// MMMMM・MMMMMMは月名と月番号を続けて表示します。
    /// ddddd・DDDDDは曜日名と日番号、dddddd・DDDDDDは曜日名と二桁の日番号を続けて表示します。
    /// hhh・HHH・mmm・sssは、二桁指定と一桁指定を続けて表示します。
    /// Wordの書式全体を.NETへ変換するものではありません。
    /// </summary>
    /// <remarks>
    /// 年はyy・yyyy（大文字のYY・YYYYも含む）を確認済みです。
    /// yyyはWordの日付選択で日時・表示が更新されなかったため、対応対象に含めません。
    /// yyy専用の変換や拒否処理は設けていません。
    /// </remarks>
    /// <param name="format">テンプレートの日付表示形式。</param>
    /// <returns>大文字の年・日指定と午前午後指定を.NETの記号へ置き換えた表示形式。</returns>
    static string ToDotNetDateFormat(string format)
    {
        // 引用部分を先に読み、その内部を書式記号として変換しません。午前午後指定はmやMへ分割しません。
        var convertedFormat = DateFormatTokens().Replace(
            format,
            it =>
            it.Value switch
            {
                "yyyyy" or "YYYYY" => "yyyy''yy",
                "hhh" => "hh''h",
                "HHH" => "HH''H",
                "mmm" => "mm''m",
                "sss" => "ss''s",
                "MMMMM" => "MMMM''M", // 空の引用文字列で月名と月番号を分け、後続のMは月番号の桁数指定へ残します。
                "ddddd" or "DDDDD" => "dddd''d", // 曜日名と日番号が一つの曜日指定へ結合されないよう区切ります。
                "DDDDDD" => "dddd''dd",
                "am/pm" or "AM/PM" => "tt",
                "/" => "'/'",
                ":" => "':'",
                "'" => "\\'",
                "\"" => "\\\"",
                "\\" => "\\\\",
                "%" => "\\%",
                "f" => "\\f",
                "F" => "\\F",
                "K" => "\\K",
                "zzz" => "'zzz'",
                "t" => "\\t",
                var token => token.StartsWith('\'')
                    ? token.Replace("\\", "\\\\")
                    : token.ToLowerInvariant(),
            });

        // 一文字書式用の%は、文書に含まれる表示文字の%をエスケープした後で付けます。
        return convertedFormat.Length == 1 ? $"%{convertedFormat}" : convertedFormat;
    }

    /// <summary>
    /// 引用文字列と、.NET用に変換するWordの日付書式記号を識別します。
    /// </summary>
    /// <returns>コンパイル時に生成される、書式記号の検索用正規表現。</returns>
    [GeneratedRegex("'[^']*'|yyyyy|MMMMM|ddddd|hhh|HHH|mmm|sss|am/pm|AM/PM|Y+|D+|/|:|'|\\\\|\"|%|f|F|K|zzz|t")]
    private static partial Regex DateFormatTokens();

    /// <summary>
    /// 日時と表示設定を保持するOOXML要素を取得します。
    /// </summary>
    Wordprocessing.SdtContentDate DateProperties =>
        Element.PropertyElements<Wordprocessing.SdtContentDate>()
            .Single();

    /// <summary>
    /// Word上で日付を表示する文字列要素を取得します。
    /// </summary>
    Wordprocessing.Text DisplayText =>
        Element.Descendants<Wordprocessing.Text>().Single();

    /// <summary>
    /// 表示文字列のrunに直接設定された言語と文字スタイル参照を取得します。
    /// </summary>
    Wordprocessing.RunProperties? DisplayRunProperties =>
        DisplayText.Ancestors<Wordprocessing.Run>().Single().RunProperties;

    /// <summary>
    /// 表示runの文字スタイル参照に対応する定義を取得します。スタイル定義がなければ参照先を探索しません。
    /// </summary>
    Wordprocessing.Style? DisplayCharacterStyle =>
        DocumentStyles?.Elements<Wordprocessing.Style>()
            .SingleOrDefault(it => it.StyleId?.Value == DisplayRunProperties?.RunStyle?.Val?.Value);

    /// <summary>
    /// 文字スタイルの言語と文書既定の言語を参照するため、所属文書のスタイル定義を取得します。
    /// </summary>
    Wordprocessing.Styles? DocumentStyles =>
        Element.Ancestors<Wordprocessing.Document>().Single()
            .MainDocumentPart?.StyleDefinitionsPart?.Styles;
}
