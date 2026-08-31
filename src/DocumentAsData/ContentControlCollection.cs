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

    internal ContentControlCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagのContent Controlを取得します。
    /// </summary>
    /// <param name="tag">取得するContent ControlのTag。</param>
    /// <returns>指定したTagのContent Control。</returns>
    public ContentControl this[string tag] =>
        (
            from contentControl in this
            where contentControl.Tag == tag
            select contentControl
        ).First();

    /// <summary>
    /// Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>Content Controlを列挙する列挙子。</returns>
    public IEnumerator<ContentControl> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            select new ContentControl(element)
        ).GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
