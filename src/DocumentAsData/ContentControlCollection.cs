using System.Collections;
using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の特定種類のContent Controlを取得するコレクションを表します。
/// </summary>
/// <typeparam name="T">Content Controlの種類。</typeparam>
public abstract class ContentControlCollection<T> : IEnumerable<T>
    where T : ContentControl
{
    /// <summary>
    /// 文書全体または明細内など、Content Controlを検索する範囲の要素列です。
    /// </summary>
    readonly IEnumerable<OpenXmlElement> elements;

    /// <summary>
    /// OOXML要素がこのコレクションで扱う種類かを判定します。
    /// </summary>
    readonly Func<Wordprocessing.SdtElement, bool> isTarget;

    /// <summary>
    /// OOXML要素に対応するContent Controlを生成します。
    /// </summary>
    readonly Func<Wordprocessing.SdtElement, T> create;

    /// <summary>
    /// 同じOOXML要素から生成したContent Controlをコレクションの生存期間中共有します。
    /// </summary>
    readonly OpenXmlElementCache<T> cache = new();

    /// <summary>
    /// 指定した範囲から、対象種類のContent Controlを取得するコレクションを作成します。
    /// </summary>
    /// <param name="elements">Content Controlを検索する範囲の要素列。</param>
    /// <param name="isTarget">対象種類のOOXML要素かを判定する処理。</param>
    /// <param name="create">OOXML要素からContent Controlを生成する処理。</param>
    private protected ContentControlCollection(
        IEnumerable<OpenXmlElement> elements,
        Func<Wordprocessing.SdtElement, bool> isTarget,
        Func<Wordprocessing.SdtElement, T> create)
    {
        this.elements = elements;
        this.isTarget = isTarget;
        this.create = create;
    }

    /// <summary>
    /// 指定したTagのContent Controlを取得します。
    /// </summary>
    /// <param name="tag">取得するContent ControlのTag。</param>
    /// <returns>指定したTagのContent Control。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagのContent Controlが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagのContent Controlが複数存在する場合。</exception>
    public T this[string tag] =>
        (
            from contentControl in this
            where contentControl.Tag == tag
            select contentControl
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// 対象種類のContent Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>Content Controlを列挙する列挙子。</returns>
    public IEnumerator<T> GetEnumerator() =>
        (
            from element in elements.OfType<Wordprocessing.SdtElement>()
            where element.HasTag && isTarget(element)
            select cache.GetOrAdd(element, () => create(element))
        ).GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}

/// <summary>
/// 文書内の文字列Content Controlを取得するコレクションを表します。
/// </summary>
public class ContentControlCollection
    : ContentControlCollection<TextContentControl>
{
    /// <summary>
    /// 文書全体を文字列Content Controlの列挙対象にします。
    /// </summary>
    /// <param name="document">文字列Content Controlを取得する文書。</param>
    internal ContentControlCollection(Document document)
        : this(document, document.Elements)
    {
    }

    /// <summary>
    /// 指定した範囲を文字列Content Controlの列挙対象にします。
    /// </summary>
    /// <param name="document">文字列Content Controlが属する文書。</param>
    /// <param name="elements">文字列Content Controlを検索する範囲の要素列。</param>
    internal ContentControlCollection(Document document, IEnumerable<OpenXmlElement> elements)
        : base(
            elements,
            element => element.IsText,
            element => new TextContentControl(document, element))
    {
    }
}
