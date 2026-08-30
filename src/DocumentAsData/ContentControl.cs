using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを表します。
/// </summary>
public class ContentControl
{
    /// <summary>
    /// このContent Controlを構成するOOXML要素です。
    /// </summary>
    readonly Wordprocessing.SdtElement element;

    internal ContentControl(Wordprocessing.SdtElement element)
    {
        this.element = element;
    }

    /// <summary>
    /// Content Controlが属する文書を取得します。
    /// </summary>
    public Document Document =>
        throw new NotImplementedException();

    /// <summary>
    /// Content ControlのTagを取得します。
    /// </summary>
    public string Tag =>
        (
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from tag in properties.Elements<Wordprocessing.Tag>()
            select tag.Val?.Value
        ).Single()!;

    /// <summary>
    /// Content Controlの値を取得または設定します。
    /// </summary>
    public string Value
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }
}
