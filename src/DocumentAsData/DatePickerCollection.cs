using System.Collections;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の日付選択Content Controlを取得するコレクションを表します。
/// </summary>
public class DatePickerCollection : IEnumerable<DatePicker>
{
    /// <summary>
    /// 日付選択Content Controlを列挙する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 同じOOXML要素から生成した日付選択Content Controlをコレクションの生存期間中共有します。
    /// </summary>
    readonly OpenXmlElementCache<DatePicker> cache = new();

    /// <summary>
    /// 文書全体を日付選択Content Controlの列挙対象にします。
    /// </summary>
    /// <param name="document">日付選択Content Controlを取得する文書。</param>
    internal DatePickerCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したTagの日付選択Content Controlを取得します。
    /// </summary>
    /// <param name="tag">取得する日付選択Content ControlのTag。</param>
    /// <returns>指定したTagの日付選択Content Control。</returns>
    /// <exception cref="KeyNotFoundException">指定したTagの日付選択Content Controlが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定したTagの日付選択Content Controlが複数存在する場合。</exception>
    public DatePicker this[string tag] =>
        (
            from datePicker in this
            where datePicker.Tag == tag
            select datePicker
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// 日付選択Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>日付選択Content Controlを列挙する列挙子。</returns>
    public IEnumerator<DatePicker> GetEnumerator() =>
        (
            from element in document.Elements.OfType<Wordprocessing.SdtElement>()
            where element.HasTag && element.IsDatePicker
            select GetOrCreateDatePicker(element)
        ).GetEnumerator();

    /// <summary>
    /// 指定したOOXML要素に対応する既存または新しい日付選択Content Controlを取得します。
    /// </summary>
    /// <param name="element">日付選択Content Controlを表すOOXML要素。</param>
    /// <returns>指定した要素に対応する日付選択Content Control。</returns>
    DatePicker GetOrCreateDatePicker(Wordprocessing.SdtElement element) =>
        cache.GetOrAdd(element, () => new(document, element));

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
