using System.Collections;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のチェックボックスを取得するコレクションを表します。
/// </summary>
public class CheckBoxCollection : IEnumerable<CheckBox>
{
    /// <summary>
    /// チェックボックスを列挙する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 同じOOXML要素から生成したチェックボックスをコレクションの生存期間中共有します。
    /// </summary>
    readonly OpenXmlElementCache<CheckBox> cache = new();

    /// <summary>
    /// 文書全体をチェックボックスの列挙対象にします。
    /// </summary>
    /// <param name="document">チェックボックスを取得する文書。</param>
    internal CheckBoxCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagのチェックボックスを取得します。
    /// </summary>
    /// <param name="tag">取得するチェックボックスのTag。</param>
    /// <returns>指定したTagのチェックボックス。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagのチェックボックスが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagのチェックボックスが複数存在する場合。</exception>
    public CheckBox this[string tag] =>
        (
            from checkBox in this
            where checkBox.Tag == tag
            select checkBox
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// チェックボックスを列挙する列挙子を返します。
    /// </summary>
    /// <returns>チェックボックスを列挙する列挙子。</returns>
    public IEnumerator<CheckBox> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            where element.HasTag && element.IsCheckBox
            select GetOrCreateCheckBox(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しいチェックボックスを取得します。
    /// </summary>
    /// <param name="element">チェックボックスを表すOOXML要素。</param>
    /// <returns>指定した要素に対応するチェックボックス。</returns>
    CheckBox GetOrCreateCheckBox(Wordprocessing.SdtElement element) =>
        cache.GetOrAdd(element, () => new(document, element));

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
