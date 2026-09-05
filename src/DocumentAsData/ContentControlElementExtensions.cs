using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// Content Controlを表すOOXML要素からプロパティ要素を取得する処理を提供します。
/// </summary>
static class ContentControlElementExtensions
{
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
