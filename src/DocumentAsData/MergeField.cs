using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを表します。
/// </summary>
public class MergeField
{
    /// <summary>
    /// 単純形式のMERGEFIELDを構成するOOXML要素です。
    /// 複合形式の場合はnullです。
    /// </summary>
    readonly Wordprocessing.SimpleField? simpleField;

    /// <summary>
    /// 複合形式の表示値を構成する文字列要素を文書順に保持します。
    /// 読み取り時はこれらを連結します。
    /// </summary>
    readonly IReadOnlyCollection<Wordprocessing.Text> complexValueTexts;

    /// <summary>
    /// OOXMLのフィールド命令から解析した名前を、単純形式と複合形式で共通に公開するため保持します。
    /// </summary>
    readonly string name;

    /// <summary>
    /// 単純フィールド要素自体を値の読み書き対象とするMERGEFIELDを作成します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="field">MERGEFIELDを構成する単純フィールド要素。</param>
    /// <param name="name">MERGEFIELDの名前。</param>
    internal MergeField(
        Document document,
        Wordprocessing.SimpleField field,
        string name)
    {
        Document = document;
        simpleField = field;
        complexValueTexts = [];
        this.name = name;
    }

    /// <summary>
    /// 開始要素と結果文字列が分かれた複合形式のMERGEFIELDを作成します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="name">MERGEFIELDの名前。</param>
    /// <param name="valueTexts">MERGEFIELDの表示値を構成する文字列要素。</param>
    internal MergeField(
        Document document,
        string name,
        IEnumerable<Wordprocessing.Text> valueTexts)
    {
        Document = document;
        complexValueTexts = valueTexts.ToArray();
        this.name = name;
    }

    /// <summary>
    /// MERGEFIELDが属する文書を取得します。
    /// </summary>
    public Document Document { get; }

    /// <summary>
    /// MERGEFIELDの名前を取得します。
    /// </summary>
    public string Name => name;

    /// <summary>
    /// MERGEFIELDの値を取得または設定します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// 値の設定時に表示値の文字列要素が存在しない場合。
    /// </exception>
    public string Value
    {
        get => simpleField?.InnerText ??
            string.Concat(complexValueTexts.Select(it => it.Text));
        set =>
            (
                simpleField?.Descendants<Wordprocessing.Text>().First() ??
                complexValueTexts.First()
            ).Text = value;
    }
}
