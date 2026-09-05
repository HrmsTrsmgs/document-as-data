using System.Collections;
using DocumentFormat.OpenXml;
using Word2010 = DocumentFormat.OpenXml.Office2010.Word;
using Word2013 = DocumentFormat.OpenXml.Office2013.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを取得するコレクションを表します。
/// </summary>
public class ContentControlCollection : IEnumerable<ContentControl>
{
    /// <summary>
    /// Content Controlを列挙する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 同じOOXML要素から生成したContent Controlをコレクションの生存期間中共有します。
    /// </summary>
    readonly OpenXmlElementCache<ContentControl> cache = new();

    /// <summary>
    /// 文書全体を列挙対象とし、列挙とTag検索で生成したContent Controlを共有できるようにします。
    /// </summary>
    /// <param name="document">Content Controlを取得する文書。</param>
    internal ContentControlCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagのContent Controlを取得します。
    /// </summary>
    /// <param name="tag">取得するContent ControlのTag。</param>
    /// <returns>指定したTagのContent Control。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagのContent Controlが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagのContent Controlが複数存在する場合。</exception>
    public ContentControl this[string tag] =>
        (
            from contentControl in this
            where contentControl.Tag == tag
            select contentControl
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>Content Controlを列挙する列挙子。</returns>
    public IEnumerator<ContentControl> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            where element.HasTag() && ContentControlType.IsText(element)
            select GetOrCreateContentControl(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しいContent Controlを取得します。
    /// </summary>
    /// <param name="element">Content Controlを表すOOXML要素。</param>
    /// <returns>指定した要素に対応するContent Control。</returns>
    ContentControl GetOrCreateContentControl(Wordprocessing.SdtElement element) =>
        cache.GetOrAdd(element, () => new(document, element));

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    /// <summary>
    /// OOXMLのContent Control種類を、文字列値を扱えるかどうかに分類します。
    /// </summary>
    static class ContentControlType
    {
        /// <summary>
        /// 指定したContent Controlが文字列値を扱う種類かを取得します。
        /// 型要素がない場合は、OOXMLの規定によりリッチテキストとして扱います。
        /// </summary>
        /// <param name="element">分類するContent ControlのOOXML要素。</param>
        /// <returns>文字列値を扱う種類の場合は<c>true</c>。</returns>
        internal static bool IsText(Wordprocessing.SdtElement element) =>
            ValueKinds(element).All(it => it == ValueKind.Text);

        /// <summary>
        /// 指定したContent Controlに明示された値の種類を列挙します。
        /// </summary>
        /// <param name="element">種類を取得するContent ControlのOOXML要素。</param>
        /// <returns>明示されている値の種類。型要素がない場合は空の列挙。</returns>
        static IEnumerable<ValueKind> ValueKinds(
            Wordprocessing.SdtElement element) =>
            from properties in element.Elements<Wordprocessing.SdtProperties>()
            from property in properties.ChildElements
            let kind = ValueKindOf(property)
            where kind != ValueKind.Unspecified
            select kind;

        /// <summary>
        /// Content Controlの種類を表すOOXML要素を、公開APIが扱う値の種類へ変換します。
        /// Tagやロックなど、Content Controlの種類を表さない要素は未指定とします。
        /// </summary>
        /// <param name="element">分類するContent ControlのOOXML要素。</param>
        /// <returns>要素が表す値の種類。</returns>
        static ValueKind ValueKindOf(OpenXmlElement element) => element switch
        {
            Wordprocessing.SdtContentText or
            Wordprocessing.SdtContentRichText => ValueKind.Text,
            Wordprocessing.SdtContentEquation or
            Wordprocessing.SdtContentPicture or
            Wordprocessing.SdtContentCitation or
            Wordprocessing.SdtContentGroup or
            Wordprocessing.SdtContentBibliography or
            Wordprocessing.SdtContentComboBox or
            Wordprocessing.SdtContentDate or
            Wordprocessing.SdtContentDocPartObject or
            Wordprocessing.SdtContentDocPartList or
            Wordprocessing.SdtContentDropDownList or
            Word2010.EntityPickerEmpty or
            Word2010.SdtContentCheckBox or
            Word2013.SdtRepeatedSection or
            Word2013.SdtRepeatedSectionItem => ValueKind.NonText,
            _ => ValueKind.Unspecified
        };

        /// <summary>
        /// Content Controlが公開APIで扱える値の種類を表します。
        /// </summary>
        enum ValueKind
        {
            /// <summary>
            /// Content Controlの値の種類を表さない要素です。
            /// </summary>
            Unspecified,

            /// <summary>
            /// 文字列として読み書きできる値です。
            /// </summary>
            Text,

            /// <summary>
            /// 種類固有の構造を持ち、文字列としては扱わない値です。
            /// </summary>
            NonText
        }
    }
}
