using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// Content Controlを表すOOXML要素からTagを扱う処理を提供します。
/// </summary>
static class ContentControlTagExtensions
{
    extension(Wordprocessing.SdtElement self)
    {
        /// <summary>
        /// このContent Control自身にTagが設定されているかを取得します。
        /// </summary>
        /// <returns>Tagが設定されている場合は<c>true</c>。</returns>
        internal bool HasTag() =>
            TagsIn(self).Any();

        /// <summary>
        /// このContent ControlのTagを取得します。
        /// </summary>
        /// <returns>Tagの値。</returns>
        /// <exception cref="InvalidOperationException">
        /// Tagの値が欠落しているか、Tagが複数存在する場合。
        /// </exception>
        internal string ReadTag() =>
            (
                from tag in TagsIn(self)
                select tag.Val?.Value
            ).Single() ?? throw new InvalidOperationException();
    }

    /// <summary>
    /// 指定したContent Control自身に設定されたTag要素を列挙します。
    /// </summary>
    /// <param name="element">Tagを探索するContent ControlのOOXML要素。</param>
    /// <returns>Content Control自身のプロパティに含まれるTag要素。</returns>
    static IEnumerable<Wordprocessing.Tag> TagsIn(
        Wordprocessing.SdtElement element) =>
        element.PropertyElements<Wordprocessing.Tag>();
}
