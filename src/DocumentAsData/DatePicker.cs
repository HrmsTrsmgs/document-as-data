using System.Globalization;
using System.Xml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の日付選択Content Controlを表します。
/// </summary>
public class DatePicker : ContentControl
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
    /// 日付選択Content Controlの日時を取得または設定します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// 日付選択のプロパティ、日時、表示形式、表示言語、表示文字列の
    /// いずれかが欠落しているか、必要な要素が複数存在する場合。
    /// </exception>
    public DateTimeOffset SelectedDateTime
    {
        get => XmlConvert.ToDateTimeOffset(
            DateProperties.FullDate?.InnerText
                ?? throw new InvalidOperationException());
        set
        {
            var properties = DateProperties;
            var fullDate = properties.FullDate
                ?? throw new InvalidOperationException();
            var format = properties.DateFormat?.Val?.Value
                ?? throw new InvalidOperationException();
            var language = properties.LanguageId?.Val?.Value
                ?? throw new InvalidOperationException();

            fullDate.InnerText = XmlConvert.ToString(value);
            DisplayText.Text = value.ToString(
                format,
                CultureInfo.GetCultureInfo(language));
        }
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
}
