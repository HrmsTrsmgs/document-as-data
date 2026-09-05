using DocumentFormat.OpenXml;

namespace Marimo.DocumentAsData;

/// <summary>
/// 同じOOXML要素から生成した公開オブジェクトを再利用します。
/// </summary>
/// <typeparam name="T">OOXML要素に対応する公開オブジェクトの型。</typeparam>
sealed class OpenXmlElementCache<T>
    where T : class
{
    /// <summary>
    /// OOXML要素を識別子として生成済みの公開オブジェクトを保持します。
    /// </summary>
    readonly Dictionary<OpenXmlElement, T> items = [];

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しい公開オブジェクトを取得します。
    /// </summary>
    /// <param name="element">公開オブジェクトを識別するOOXML要素。</param>
    /// <param name="create">未生成の場合に公開オブジェクトを作成する処理。</param>
    /// <returns>既存または新しく生成した公開オブジェクト。</returns>
    internal T GetOrAdd(OpenXmlElement element, Func<T> create)
    {
        if (items.TryGetValue(element, out var item))
        {
            return item;
        }

        item = create();
        items.Add(element, item);
        return item;
    }
}
