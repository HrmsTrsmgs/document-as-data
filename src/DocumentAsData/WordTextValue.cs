using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 公開APIの文字列とWordprocessingMLの文字列、タブ、改行要素を相互変換します。
/// </summary>
static class WordTextValue
{
    /// <summary>
    /// OOXMLの文字列、タブ、改行を公開APIの文字列表現へ戻します。
    /// </summary>
    /// <param name="elements">値を構成するOOXML要素。</param>
    /// <returns>タブをタブ文字、改行をCRLFで表した値。</returns>
    internal static string Read(IEnumerable<OpenXmlElement> elements) =>
        string.Concat(
            from element in elements
            where element is Wordprocessing.Text or
                Wordprocessing.TabChar or
                Wordprocessing.Break
            select element switch
            {
                Wordprocessing.Text text => text.Text,
                Wordprocessing.TabChar => "\t",
                Wordprocessing.Break => "\r\n",
                _ => throw new InvalidOperationException()
            });

    /// <summary>
    /// 公開APIの文字列をWordが表示できる文字列、タブ、改行要素へ分解します。
    /// </summary>
    /// <param name="value">要素へ変換する値。</param>
    /// <returns>文書順に並んだOOXML要素。</returns>
    internal static IEnumerable<OpenXmlElement> CreateElements(string value) =>
        from part in Regex.Split(value, "(\r\n|\n|\t)")
        where !part.IsEmpty()
        select CreateElement(part);

    /// <summary>
    /// 分解済みの文字列片を対応するOOXML要素へ変換します。
    /// </summary>
    /// <param name="part">通常文字列、タブ、CRLF、またはLF。</param>
    /// <returns>文字列片に対応するOOXML要素。</returns>
    static OpenXmlElement CreateElement(string part) => part switch
    {
        "\t" => new Wordprocessing.TabChar(),
        "\r\n" or "\n" => new Wordprocessing.Break(),
        _ => new Wordprocessing.Text(part)
    };
}
