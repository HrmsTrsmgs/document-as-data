using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを表します。
/// </summary>
public class ContentControl
{
    /// <summary>
    /// このContent Controlが属する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// このContent Controlを構成するOOXML要素です。
    /// </summary>
    readonly Wordprocessing.SdtElement element;

    /// <summary>
    /// 文書内のOOXML要素への参照を保持し、Valueの読み書きを同じ要素へ反映できるようにします。
    /// </summary>
    /// <param name="document">Content Controlが属する文書。</param>
    /// <param name="element">Content Controlを構成するOOXML要素。</param>
    internal ContentControl(Document document, Wordprocessing.SdtElement element)
    {
        this.document = document;
        this.element = element;
    }

    /// <summary>
    /// Content Controlが属する文書を取得します。
    /// </summary>
    public Document Document =>
        document;

    /// <summary>
    /// Content ControlのTagを取得します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Tagの値が欠落しているか、Tagが複数存在する場合。
    /// </exception>
    public string Tag =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from tag in properties.Elements<Wordprocessing.Tag>()
            select tag.Val?.Value
        ).Single() ?? throw new InvalidOperationException();

    /// <summary>
    /// Content Controlの値を取得または設定します。
    /// </summary>
    /// <remarks>
    /// プレースホルダー表示中は、表示用文字列ではなく空文字列を返します。
    /// 値を設定すると、プレースホルダー表示状態を解除します。
    /// </remarks>
    /// <exception cref="InvalidOperationException">
    /// 値の設定時に文字列要素が存在しない場合。
    /// </exception>
    public string Value
    {
        get => IsShowingPlaceholder ? "" :
            ReadValue(element.Descendants());
        set
        {
            var texts = element.Descendants<Wordprocessing.Text>().ToArray();

            RemovePlaceholderState();
            ReplaceFirstText(texts.First(), value);
            Array.ForEach(texts[1..], it => it.Text = "");
        }
    }

    /// <summary>
    /// OOXMLの文字列、タブ、改行を公開APIの文字列表現へ戻します。
    /// </summary>
    /// <param name="elements">値を構成するOOXML要素。</param>
    /// <returns>タブをタブ文字、改行をCRLFで表した値。</returns>
    static string ReadValue(IEnumerable<OpenXmlElement> elements) =>
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
    /// 最初の文字列要素を、Wordが表示できる文字列、タブ、改行要素へ置き換えます。
    /// </summary>
    /// <param name="text">置換対象の最初の文字列要素。</param>
    /// <param name="value">設定する値。</param>
    static void ReplaceFirstText(Wordprocessing.Text text, string value)
    {
        var parent = text.Parent ?? throw new InvalidOperationException();

        text.Remove();
        parent.Append(CreateValueElements(value));
    }

    /// <summary>
    /// 公開APIの文字列をWordが表示できる文字列、タブ、改行要素へ分解します。
    /// </summary>
    /// <param name="value">要素へ変換する値。</param>
    /// <returns>文書順に並んだOOXML要素。</returns>
    static IEnumerable<OpenXmlElement> CreateValueElements(string value) =>
        from part in Regex.Split(value, "(\r\n|\t)")
        where part.Length > 0
        select CreateValueElement(part);

    /// <summary>
    /// 分解済みの文字列片を対応するOOXML要素へ変換します。
    /// </summary>
    /// <param name="part">通常文字列、タブ、またはCRLF。</param>
    /// <returns>文字列片に対応するOOXML要素。</returns>
    static OpenXmlElement CreateValueElement(string part) => part switch
    {
        "\t" => new Wordprocessing.TabChar(),
        "\r\n" => new Wordprocessing.Break(),
        _ => new Wordprocessing.Text(part)
    };

    /// <summary>
    /// このContent Controlがプレースホルダー表示中かを取得します。
    /// </summary>
    bool IsShowingPlaceholder =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from placeholder in properties.Elements<Wordprocessing.ShowingPlaceholder>()
            select placeholder
        ).Any();

    /// <summary>
    /// 値の書き込み後に表示文字列がプレースホルダーとして扱われないよう、表示状態を解除します。
    /// </summary>
    void RemovePlaceholderState() =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from placeholder in properties.Elements<Wordprocessing.ShowingPlaceholder>()
            select placeholder
        ).SingleOrDefault()?.Remove();
}
