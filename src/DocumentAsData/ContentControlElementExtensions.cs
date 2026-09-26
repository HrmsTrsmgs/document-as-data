using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// Content ControlのOOXMLプロパティと、要素の所属範囲を取得する処理を提供します。
/// </summary>
static class ContentControlElementExtensions
{
    extension(OpenXmlElement self)
    {
        /// <summary>
        /// この要素を囲むContent Controlのいずれかが繰り返しセクションかを取得します。
        /// 要素自身は含めず、通常項目と明細内の項目を区別します。
        /// </summary>
        internal bool IsInRepeatingSection =>
            self.Ancestors<Wordprocessing.SdtElement>().Any(it => it.IsRepeatingSection);
    }

    extension(Wordprocessing.SdtElement self)
    {
        /// <summary>
        /// このContent Control自身のプロパティに含まれる指定型の要素を列挙します。
        /// </summary>
        /// <typeparam name="T">取得するプロパティ要素の型。</typeparam>
        /// <returns>Content Control自身のプロパティに含まれる指定型の要素。</returns>
        internal IEnumerable<T> PropertyElements<T>()
            where T : OpenXmlElement =>
            from properties in self.Elements<Wordprocessing.SdtProperties>()
            from property in properties.Elements<T>()
            select property;
    }
}
