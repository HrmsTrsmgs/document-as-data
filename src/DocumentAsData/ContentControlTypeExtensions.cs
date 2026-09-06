using DocumentFormat.OpenXml;
using Word2010 = DocumentFormat.OpenXml.Office2010.Word;
using Word2013 = DocumentFormat.OpenXml.Office2013.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// OOXMLのContent Control種類を、DocumentAsDataが扱う値の種類へ分類する拡張プロパティを提供します。
/// </summary>
static class ContentControlTypeExtensions
{
    extension(Wordprocessing.SdtElement self)
    {
        /// <summary>
        /// このContent Controlが文字列値を扱う種類かを取得します。
        /// 型要素がない場合は、OOXMLの規定によりリッチテキストとして扱います。
        /// </summary>
        internal bool IsText =>
            ValueKinds(self).All(it => it == ValueKind.Text);

        /// <summary>
        /// このContent Controlがチェックボックスかを取得します。
        /// </summary>
        internal bool IsCheckBox =>
            self.PropertyElements<Word2010.SdtContentCheckBox>().Any();

        /// <summary>
        /// このContent Controlが日付選択かを取得します。
        /// </summary>
        internal bool IsDatePicker =>
            self.PropertyElements<Wordprocessing.SdtContentDate>().Any();
    }

    /// <summary>
    /// 指定したContent Controlに明示された値の種類を列挙します。
    /// </summary>
    /// <param name="element">種類を取得するContent ControlのOOXML要素。</param>
    /// <returns>明示されている値の種類。型要素がない場合は空の列挙。</returns>
    static IEnumerable<ValueKind> ValueKinds(
        Wordprocessing.SdtElement element) =>
        from property in element.PropertyElements<OpenXmlElement>()
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
