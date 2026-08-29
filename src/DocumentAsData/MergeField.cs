using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを表します。
/// </summary>
public class MergeField
{
    readonly Wordprocessing.SimpleField field;
    readonly string name;

    internal MergeField(
        Document document,
        Wordprocessing.SimpleField field,
        string name)
    {
        Document = document;
        this.field = field;
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
        get => this.field.InnerText;
        set => throw new NotImplementedException();
    }
}
