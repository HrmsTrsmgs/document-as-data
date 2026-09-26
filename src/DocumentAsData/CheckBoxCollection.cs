using DocumentFormat.OpenXml;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のチェックボックスを取得するコレクションを表します。
/// </summary>
public class CheckBoxCollection : ContentControlCollection<CheckBox>
{
    /// <summary>
    /// 文書全体をチェックボックスの列挙対象にします。
    /// </summary>
    /// <param name="document">チェックボックスを取得する文書。</param>
    internal CheckBoxCollection(Document document)
        : this(document, document.Elements)
    {
    }

    /// <summary>
    /// 指定した範囲をチェックボックスの列挙対象にします。
    /// </summary>
    /// <param name="document">チェックボックスが属する文書。</param>
    /// <param name="elements">チェックボックスを検索する範囲の要素列。</param>
    internal CheckBoxCollection(Document document, IEnumerable<OpenXmlElement> elements)
        : base(
            elements,
            element => element.IsCheckBox,
            element => new CheckBox(document, element))
    {
    }
}
