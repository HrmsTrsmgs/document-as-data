using System.Globalization;
using System.Text.RegularExpressions;
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
    /// 種類と、引用符で囲んだTagを返します。
    /// </summary>
    /// <returns>種類とTagを確認できる表示文字列。</returns>
    public override string ToString() =>
        $"DatePicker {{ Tag = {QuoteName(Tag)} }}";

    /// <summary>
    /// 日付選択Content Controlの日時を取得または設定します。
    /// </summary>
    /// <remarks>
    /// 日時が未入力でも設定できます。表示形式、表示言語、表示文字列は必要です。
    /// </remarks>
    /// <exception cref="InvalidOperationException">
    /// 取得時に日時が欠落している場合。または日付選択のプロパティや、設定時に必要な
    /// 表示形式、表示言語、表示文字列が欠落しているか、必要な要素が複数存在する場合。
    /// </exception>
    public DateTimeOffset SelectedDateTime
    {
        get => XmlConvert.ToDateTimeOffset(
            DateProperties.FullDate?.InnerText
                ?? throw new InvalidOperationException());
        set
        {
            var properties = DateProperties;
            var format = properties.DateFormat?.Val?.Value
                ?? throw new InvalidOperationException();
            var language = properties.LanguageId?.Val?.Value
                ?? throw new InvalidOperationException();

            properties.FullDate = new() { InnerText = XmlConvert.ToString(value) };
            DisplayText.Text = value.ToString(
                ToDotNetDateFormat(format),
                CultureInfo.GetCultureInfo(language));
        }
    }

    /// <summary>
    /// 単一引用符で囲まれた表示文字列を保持し、Wordの年・日・午前午後指定を.NET用に変換します。
    /// 一文字の書式も、.NETの標準書式ではなく指定された一項目として扱います。
    /// スラッシュは地域別の区切り文字へ変更せず、そのまま表示します。
    /// Wordの書式全体を.NETへ変換するものではありません。
    /// </summary>
    /// <param name="format">テンプレートの日付表示形式。</param>
    /// <returns>大文字の年・日指定と午前午後指定を.NETの記号へ置き換えた表示形式。</returns>
    static string ToDotNetDateFormat(string format) =>
        // 引用部分を先に読み、その内部を変換しません。午前午後指定はmやMへ分割しません。
        Regex.Replace(
            format.Length == 1 ? $"%{format}" : format,
            "'[^']*'|am/pm|AM/PM|Y+|D+|/", it =>
            it.Value switch
            {
                "am/pm" or "AM/PM" => "tt",
                "/" => "'/'",
                var token => token.StartsWith('\'') ? token : token.ToLowerInvariant(),
            });

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
