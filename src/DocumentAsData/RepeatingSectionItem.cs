using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 繰り返しセクション内の明細1件を表します。
/// </summary>
public class RepeatingSectionItem : ContentControl
{
    /// <summary>
    /// 明細1件を囲むOOXML要素への参照を保持します。
    /// </summary>
    /// <param name="document">明細が属する文書。</param>
    /// <param name="element">明細1件を構成するOOXML要素。</param>
    internal RepeatingSectionItem(Document document, Wordprocessing.SdtElement element)
        : base(document, element)
    {
        ContentControls = new(document, element.Descendants());
        MergeFields = new(document, element.Descendants());
    }

    /// <summary>
    /// この明細内の文字列Content Controlを取得します。
    /// </summary>
    public ContentControlCollection ContentControls { get; }

    /// <summary>
    /// この明細内のMERGEFIELDを取得します。
    /// </summary>
    public MergeFieldCollection MergeFields { get; }

    /// <summary>
    /// この明細内のチェックボックスを取得します。
    /// </summary>
    public CheckBoxCollection CheckBoxes =>
        throw new NotImplementedException();

    /// <summary>
    /// この明細内の日付選択Content Controlを取得します。
    /// </summary>
    public DatePickerCollection DatePickers =>
        throw new NotImplementedException();
}
