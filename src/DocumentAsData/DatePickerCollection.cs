namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の日付選択Content Controlを取得するコレクションを表します。
/// </summary>
public class DatePickerCollection : ContentControlCollection<DatePicker>
{
    /// <summary>
    /// 文書全体を日付選択Content Controlの列挙対象にします。
    /// </summary>
    /// <param name="document">日付選択Content Controlを取得する文書。</param>
    internal DatePickerCollection(Document document)
        : base(
            document,
            element => element.IsDatePicker,
            element => new DatePicker(document, element))
    {
    }
}
