namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の繰り返しセクションを取得するコレクションを表します。
/// </summary>
public class RepeatingSectionCollection : ContentControlCollection<RepeatingSection>
{
    /// <summary>
    /// 文書全体を繰り返しセクションの列挙対象にします。
    /// </summary>
    /// <param name="document">繰り返しセクションを取得する文書。</param>
    internal RepeatingSectionCollection(Document document)
        : base(
            document.Elements,
            element => element.IsRepeatingSection,
            element => new RepeatingSection(document, element))
    {
    }
}
