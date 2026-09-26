using Word2013 = DocumentFormat.OpenXml.Office2013.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 複数件の内容をひとまとまりとして扱う、文書内の繰り返しセクションを表します。
/// </summary>
public class RepeatingSection : ContentControl
{
    /// <summary>
    /// 繰り返し全体を囲むOOXML要素への参照を保持します。
    /// </summary>
    /// <param name="document">繰り返しセクションが属する文書。</param>
    /// <param name="element">繰り返しセクションを構成するOOXML要素。</param>
    internal RepeatingSection(
        Document document,
        Wordprocessing.SdtElement element)
        : base(document, element)
    {
    }

    /// <summary>
    /// セクション内の明細を文書に現れる順で取得します。
    /// </summary>
    public IReadOnlyList<RepeatingSectionItem> Items =>
    [
        // w15:repeatingSectionItemは明細1件の枠を示します。
        // その内側にある通常の入力項目は明細として数えません。
        .. from element in Element.Descendants<Wordprocessing.SdtElement>()
           where element.PropertyElements<Word2013.SdtRepeatedSectionItem>().Any()
           select new RepeatingSectionItem(Document, element)
    ];

    /// <summary>
    /// 各明細の名前付き項目を、指定した型へ対応付けて読み込みます。
    /// </summary>
    /// <typeparam name="T">明細1件分のデータを読み込む型。</typeparam>
    /// <returns>文書順に読み込んだ明細データ。</returns>
    public IEnumerable<T> Read<T>() =>
        from item in Items
        select item.Read<T>();

    /// <summary>
    /// 既存の明細の内容を、同じ件数のデータで順番に置き換えます。
    /// </summary>
    /// <typeparam name="T">明細1件分のデータを保持する型。</typeparam>
    /// <param name="items">既存の明細と同じ件数の、置き換え後のデータ。</param>
    public void Replace<T>(IEnumerable<T> items)
    {
        foreach (var (item, data) in Items.Zip(items))
        {
            item.Replace(data);
        }
    }
}
