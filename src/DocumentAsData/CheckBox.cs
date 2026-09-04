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
}
