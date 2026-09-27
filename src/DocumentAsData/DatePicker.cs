using System.Globalization;
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
                DisplayFormat.Convert(format),
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
