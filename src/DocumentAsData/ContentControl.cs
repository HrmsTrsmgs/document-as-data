using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを表します。
/// </summary>
public abstract class ContentControl : DocumentItem
{
    /// <summary>
    /// 各種Content Controlが値を読み書きするOOXML要素を取得します。
    /// 公開APIにOpen XML SDKの型を露出させないよう、同じアセンブリの派生型だけが利用できます。
    /// </summary>
    private protected Wordprocessing.SdtElement Element { get; }

    /// <summary>
    /// 文書内のOOXML要素を保持するContent Controlを作成します。
    /// </summary>
    /// <param name="document">Content Controlが属する文書。</param>
    /// <param name="element">Content Controlを構成するOOXML要素。</param>
    private protected ContentControl(
        Document document,
        Wordprocessing.SdtElement element)
        : base(document)
    {
        Element = element;
    }

    /// <summary>
    /// Content ControlのTagを取得します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Tagの値が欠落しているか、Tagが複数存在する場合。
    /// </exception>
    public string Tag =>
        Element.Tag;

    /// <summary>
    /// Custom XMLに連結された項目では、表示値だけを読み書きしても連結先と同期できないため拒否します。
    /// 親の繰り返しセクションが連結されている場合も、Wordが明細を作り直すため対象に含めます。
    /// </summary>
    /// <exception cref="NotSupportedException">Custom XMLへのデータ連結がある場合。</exception>
    private protected void EnsureNotDataBound()
    {
        if (
            (
                from element in Element.Ancestors<Wordprocessing.SdtElement>().Prepend(Element)
                where element == Element || element.IsRepeatingSection
                from binding in element.PropertyElements<Wordprocessing.DataBinding>()
                select binding
            ).Any())
        {
            throw new NotSupportedException();
        }
    }
}
