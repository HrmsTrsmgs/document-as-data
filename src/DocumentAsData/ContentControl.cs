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
    public string Tag =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from tag in properties.Elements<Wordprocessing.Tag>()
            select tag.Val?.Value
        ).Single() ?? throw new InvalidOperationException();

    /// <summary>
    /// Content Controlの値を取得または設定します。
    /// </summary>
    public string Value
    {
        get =>
            string.Concat(
                from text in element.Descendants<Wordprocessing.Text>()
                select text.Text);
        set =>
            (
                from text in element.Descendants<Wordprocessing.Text>()
                select text
            ).Single().Text = value;
    }
}
