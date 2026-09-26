using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 複数件の内容をひとまとまりとして扱う、文書内の繰り返しセクションを表します。
/// </summary>
public class RepeatingSection : ContentControl
{
    /// <summary>
    /// 繰り返し全体を囲むOOXML要素への参照を保持します。
    /// </summary>
    /// <param name="document">繰り返しセクションが属する文書。</param>
    /// <param name="element">繰り返しセクションを構成するOOXML要素。</param>
    internal RepeatingSection(
        Document document,
        Wordprocessing.SdtElement element)
        : base(document, element)
    {
    }
}
