using System.Collections;
using Word2010 = DocumentFormat.OpenXml.Office2010.Word;
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
            where IsCheckBox(element)
            select new CheckBox(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したContent Controlがチェックボックスかを取得します。
    /// </summary>
    /// <param name="element">確認するContent ControlのOOXML要素。</param>
    /// <returns>チェックボックスの場合は<c>true</c>。</returns>
    static bool IsCheckBox(Wordprocessing.SdtElement element) =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from checkBox in properties.Elements<Word2010.SdtContentCheckBox>()
            select checkBox
        ).Any();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
