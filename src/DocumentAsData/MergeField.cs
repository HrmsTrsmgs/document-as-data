using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを表します。
/// </summary>
public class MergeField
{
    readonly Wordprocessing.SimpleField? simpleField;
    readonly IReadOnlyCollection<Wordprocessing.Text> complexValueTexts;
    readonly string name;

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
    public string Value
    {
        get => simpleField?.InnerText ??
            string.Concat(complexValueTexts.Select(it => it.Text));
        set
        {
            if (simpleField is null)
            {
                throw new NotImplementedException();
            }

            simpleField.Descendants<Wordprocessing.Text>().First().Text = value;
        }
    }
}
