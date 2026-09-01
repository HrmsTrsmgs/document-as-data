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
    /// 複合形式の命令部分と表示結果を区切る要素です。
    /// 単純形式の場合はnullです。
    /// </summary>
    readonly Wordprocessing.FieldChar? complexResultSeparator;

    /// <summary>
    /// 複合形式の終了要素です。
    /// 単純形式の場合はnullです。
    /// </summary>
    readonly Wordprocessing.FieldChar? complexFieldEnd;

    /// <summary>
    /// 複合形式の表示値を構成する文字列要素を文書順に保持します。
    /// 読み取り時はこれらを連結し、書き込み時は置換後の要素へ更新します。
    /// </summary>
    readonly List<Wordprocessing.Text> complexValueTexts;

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
        complexResultSeparator = null;
        complexFieldEnd = null;
        complexValueTexts = [];
        this.name = name;
    }

    /// <summary>
    /// 開始要素と結果文字列が分かれた複合形式のMERGEFIELDを作成します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="name">MERGEFIELDの名前。</param>
    /// <param name="resultSeparator">命令部分と表示結果を区切る要素。</param>
    /// <param name="fieldEnd">複合フィールドの終了要素。</param>
    /// <param name="valueTexts">MERGEFIELDの表示値を構成する文字列要素。</param>
    internal MergeField(
        Document document,
        string name,
        Wordprocessing.FieldChar resultSeparator,
        Wordprocessing.FieldChar fieldEnd,
        IEnumerable<Wordprocessing.Text> valueTexts)
    {
        Document = document;
        complexResultSeparator = resultSeparator;
        complexFieldEnd = fieldEnd;
        complexValueTexts = valueTexts.ToList();
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
    /// 値の設定時に表示結果が存在しないか、複合フィールドの結果領域を取得できない場合。
    /// </exception>
    public string Value
    {
        get => simpleField?.InnerText ??
            string.Concat(complexValueTexts.Select(it => it.Text));
        set
        {
            if (simpleField is not null)
            {
                simpleField.Descendants<Wordprocessing.Text>().First().Text = value;
                return;
            }

            SetComplexValue(value);
        }
    }

    /// <summary>
    /// 複合フィールドの古い表示結果を除去し、新しい値を持つ結果だけに置き換えます。
    /// </summary>
    /// <param name="value">設定する値。</param>
    /// <exception cref="InvalidOperationException">
    /// 表示結果が存在しないか、結果領域の境界を取得できない場合。
    /// </exception>
    void SetComplexValue(string value)
    {
        if (complexValueTexts.Count == 0)
        {
            throw new InvalidOperationException();
        }

        var resultStart = complexResultSeparator?.Parent ??
            throw new InvalidOperationException();
        var resultEnd = complexFieldEnd?.Parent ??
            throw new InvalidOperationException();
        if (resultStart.Parent != resultEnd.Parent)
        {
            throw new InvalidOperationException();
        }

        var oldResult = resultStart.ElementsAfter()
            .TakeWhile(it => it != resultEnd)
            .ToArray();
        Array.ForEach(oldResult, it => it.Remove());

        var text = new Wordprocessing.Text(value);
        resultEnd.InsertBeforeSelf(new Wordprocessing.Run(text));
        complexValueTexts.Clear();
        complexValueTexts.Add(text);
    }
}
