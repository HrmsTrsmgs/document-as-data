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
        : base(
            document,
            element => element.IsCheckBox,
            element => new CheckBox(document, element))
    {
    }
}
