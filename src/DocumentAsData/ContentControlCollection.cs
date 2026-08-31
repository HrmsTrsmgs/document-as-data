using System.Collections;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを取得するコレクションを表します。
/// </summary>
public class ContentControlCollection : IEnumerable<ContentControl>
{
    /// <summary>
    /// Content Controlを列挙する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 同じOOXML要素から生成したContent Controlをコレクションの生存期間中共有します。
    /// </summary>
    readonly Dictionary<Wordprocessing.SdtElement, ContentControl> cache = [];

    internal ContentControlCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagのContent Controlを取得します。
    /// </summary>
    /// <param name="tag">取得するContent ControlのTag。</param>
    /// <returns>指定したTagのContent Control。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagのContent Controlが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagのContent Controlが複数存在する場合。</exception>
    public ContentControl this[string tag] =>
        (
            from contentControl in this
            where contentControl.Tag == tag
            select contentControl
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>Content Controlを列挙する列挙子。</returns>
    public IEnumerator<ContentControl> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            select GetOrCreateContentControl(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しいContent Controlを取得します。
    /// </summary>
    /// <param name="element">Content Controlを表すOOXML要素。</param>
    /// <returns>指定した要素に対応するContent Control。</returns>
    ContentControl GetOrCreateContentControl(Wordprocessing.SdtElement element)
    {
        if (cache.TryGetValue(element, out var contentControl))
        {
            return contentControl;
        }

        contentControl = new(element);
        cache.Add(element, contentControl);
        return contentControl;
    }

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
