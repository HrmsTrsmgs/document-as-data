using Word2010 = DocumentFormat.OpenXml.Office2010.Word;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のチェックボックスを表します。
/// </summary>
public class CheckBox
{
    /// <summary>
    /// このチェックボックスが属する文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// このチェックボックスを構成するOOXML要素です。
    /// </summary>
    readonly Wordprocessing.SdtElement element;

    /// <summary>
    /// 文書内のOOXML要素への参照を保持し、チェックボックスの情報を同じ要素から取得できるようにします。
    /// </summary>
    /// <param name="document">チェックボックスが属する文書。</param>
    /// <param name="element">チェックボックスを構成するOOXML要素。</param>
    internal CheckBox(Document document, Wordprocessing.SdtElement element)
    {
        this.document = document;
        this.element = element;
    }

    /// <summary>
    /// チェックボックスが属する文書を取得します。
    /// </summary>
    public Document Document =>
        document;

    /// <summary>
    /// チェックボックスのTagを取得します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Tagの値が欠落しているか、Tagが複数存在する場合。
    /// </exception>
    public string Tag =>
        element.ReadTag();

    /// <summary>
    /// チェックボックスがチェックされているかを取得または設定します。
    /// </summary>
    /// <remarks>
    /// 設定時は、文書に定義されたチェック状態の表示文字も更新します。
    /// </remarks>
    public bool IsChecked
    {
        get => IsCheckedValue(CheckedValue.Val?.Value);
        set
        {
            CheckedValue.Val = value ?
                Word2010.OnOffValues.One :
                Word2010.OnOffValues.Zero;
            DisplayText.Text = DisplayCharacter(value);
        }
    }

    /// <summary>
    /// このチェックボックスの種類と表示文字を保持するOOXML要素を取得します。
    /// </summary>
    Word2010.SdtContentCheckBox CheckBoxProperties =>
        element.PropertyElements<Word2010.SdtContentCheckBox>().Single();

    /// <summary>
    /// このチェックボックスの状態を保持するOOXML要素を取得します。
    /// </summary>
    Word2010.Checked CheckedValue =>
        CheckBoxProperties.Checked ?? throw new InvalidOperationException();

    /// <summary>
    /// Word上でチェックボックスを表示する文字列要素を取得します。
    /// </summary>
    Wordprocessing.Text DisplayText =>
        element.Descendants<Wordprocessing.Text>().Single();

    /// <summary>
    /// 指定したチェック状態に対して文書に設定された表示文字を取得します。
    /// </summary>
    /// <param name="isChecked">表示するチェック状態。</param>
    /// <returns>チェック状態に対応する表示文字。</returns>
    string DisplayCharacter(bool isChecked) =>
        CharacterFromHexadecimal(
            isChecked ?
                CheckBoxProperties.CheckedState?.Val?.Value :
                CheckBoxProperties.UncheckedState?.Val?.Value);

    /// <summary>
    /// OOXMLに16進数で保存されたUnicodeコードポイントを文字列へ変換します。
    /// </summary>
    /// <param name="value">16進数のUnicodeコードポイント。</param>
    /// <returns>コードポイントが表す文字。</returns>
    static string CharacterFromHexadecimal(string? value) =>
        char.ConvertFromUtf32(
            Convert.ToInt32(
                value ?? throw new InvalidOperationException(),
                16));

    /// <summary>
    /// OOXMLのオン・オフ値がチェック済みを表すかを取得します。
    /// </summary>
    /// <param name="value">確認するOOXMLのオン・オフ値。</param>
    /// <returns><c>true</c>または<c>1</c>を表す場合は<c>true</c>。</returns>
    static bool IsCheckedValue(Word2010.OnOffValues? value) =>
        value == Word2010.OnOffValues.True ||
        value == Word2010.OnOffValues.One;
}
