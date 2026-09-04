using Word2010 = DocumentFormat.OpenXml.Office2010.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のチェックボックスを表します。
/// </summary>
public class CheckBox
{
    /// <summary>
    /// このチェックボックスを構成するOOXML要素です。
    /// </summary>
    readonly Wordprocessing.SdtElement element;

    /// <summary>
    /// 文書内のOOXML要素への参照を保持し、チェックボックスの情報を同じ要素から取得できるようにします。
    /// </summary>
    /// <param name="element">チェックボックスを構成するOOXML要素。</param>
    internal CheckBox(Wordprocessing.SdtElement element)
    {
        this.element = element;
    }

    /// <summary>
    /// チェックボックスのTagを取得します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Tagの値が欠落しているか、Tagが複数存在する場合。
    /// </exception>
    public string Tag =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from tag in properties.Elements<Wordprocessing.Tag>()
            select tag.Val?.Value
        ).Single() ?? throw new InvalidOperationException();

    /// <summary>
    /// チェックボックスがチェックされているかを取得します。
    /// </summary>
    public bool IsChecked =>
        IsCheckedValue(
            (
                from properties in element.Elements<Wordprocessing.SdtProperties>()
                from checkBox in properties.Elements<Word2010.SdtContentCheckBox>()
                from checkedValue in checkBox.Elements<Word2010.Checked>()
                select checkedValue.Val?.Value
            ).Single());

    /// <summary>
    /// OOXMLのオン・オフ値がチェック済みを表すかを取得します。
    /// </summary>
    /// <param name="value">確認するOOXMLのオン・オフ値。</param>
    /// <returns><c>true</c>または<c>1</c>を表す場合は<c>true</c>。</returns>
    static bool IsCheckedValue(Word2010.OnOffValues? value) =>
        value == Word2010.OnOffValues.True ||
        value == Word2010.OnOffValues.One;
}
