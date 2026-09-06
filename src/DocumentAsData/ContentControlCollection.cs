using System.Collections;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の文字列Content Controlを取得するコレクションを表します。
/// </summary>
public class ContentControlCollection : IEnumerable<TextContentControl>
{
    /// <summary>
    /// 文字列Content Controlを列挙する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 同じOOXML要素から生成した文字列Content Controlをコレクションの生存期間中共有します。
    /// </summary>
    readonly OpenXmlElementCache<TextContentControl> cache = new();

    /// <summary>
    /// 文書全体を列挙対象とし、列挙とTag検索で生成したContent Controlを共有できるようにします。
    /// </summary>
    /// <param name="document">文字列Content Controlを取得する文書。</param>
    internal ContentControlCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagの文字列Content Controlを取得します。
    /// </summary>
    /// <param name="tag">取得する文字列Content ControlのTag。</param>
    /// <returns>指定したTagの文字列Content Control。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagの文字列Content Controlが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagの文字列Content Controlが複数存在する場合。</exception>
    public TextContentControl this[string tag] =>
        (
            from contentControl in this
            where contentControl.Tag == tag
            select contentControl
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// 文字列Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>文字列Content Controlを列挙する列挙子。</returns>
    public IEnumerator<TextContentControl> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            where element.HasTag && element.IsText
            select GetOrCreateTextContentControl(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しいContent Controlを取得します。
    /// </summary>
    /// <param name="element">文字列Content Controlを表すOOXML要素。</param>
    /// <returns>指定した要素に対応する文字列Content Control。</returns>
    TextContentControl GetOrCreateTextContentControl(
        Wordprocessing.SdtElement element) =>
        cache.GetOrAdd(element, () => new(document, element));

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
