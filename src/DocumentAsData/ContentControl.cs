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
            string.Concat(
                from text in element.Descendants<Wordprocessing.Text>()
                select text.Text);
        set
        {
            var texts = element.Descendants<Wordprocessing.Text>().ToArray();

            RemovePlaceholderState();
            texts.First().Text = value;
            Array.ForEach(texts[1..], it => it.Text = "");
        }
    }

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
        Array.ForEach(
            (
                from properties in element.Elements<Wordprocessing.SdtProperties>()
                from placeholder in properties.Elements<Wordprocessing.ShowingPlaceholder>()
                select placeholder
            ).ToArray(),
            it => it.Remove());
}
