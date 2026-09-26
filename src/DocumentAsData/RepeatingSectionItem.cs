using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 繰り返しセクション内の明細1件を表します。
/// </summary>
public class RepeatingSectionItem : ContentControl
{
    /// <summary>
    /// この明細内の項目を自作クラスへ対応付ける処理です。
    /// </summary>
    readonly DocumentObjectMapper objectMapper;

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
        CheckBoxes = new(document, element.Descendants());
        DatePickers = new(document, element.Descendants());
        objectMapper = new(ContentControls, MergeFields, CheckBoxes, DatePickers);
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
    public CheckBoxCollection CheckBoxes { get; }

    /// <summary>
    /// この明細内の日付選択Content Controlを取得します。
    /// </summary>
    public DatePickerCollection DatePickers { get; }

    /// <summary>
    /// この明細内だけを対象として、既存のオブジェクト対応付けで読み取ります。
    /// </summary>
    /// <typeparam name="T">明細のデータを読み込む型。</typeparam>
    /// <returns>明細のデータを読み込んだオブジェクト。</returns>
    internal T Read<T>() =>
        objectMapper.Read<T>();

    /// <summary>
    /// この明細内だけを対象として、既存のオブジェクト対応付けで書き込みます。
    /// </summary>
    /// <typeparam name="T">明細のデータを保持する型。</typeparam>
    /// <param name="data">明細へ書き込むデータ。</param>
    internal void Replace<T>(T data) =>
        objectMapper.Replace(data);

    /// <summary>
    /// w15:repeatingSectionItemを持つ明細の枠を内容ごと複製し、同じ親の直後へ挿入します。
    /// 内側の入力項目だけでなく明細の枠も残すことで、追加分を明細として取得できます。
    /// </summary>
    internal void InsertCopyAfter() =>
        Element.InsertAfterSelf(Element.CloneNode(true));
}
