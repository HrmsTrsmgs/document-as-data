namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを表します。
/// </summary>
public class MergeField
{
    readonly string name;

    internal MergeField(string name)
    {
        this.name = name;
    }

    /// <summary>
    /// MERGEFIELDが属する文書を取得します。
    /// </summary>
    public Document Document =>
        throw new NotImplementedException();

    /// <summary>
    /// MERGEFIELDの名前を取得します。
    /// </summary>
    public string Name => name;

    /// <summary>
    /// MERGEFIELDの値を取得または設定します。
    /// </summary>
    public string Value
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }
}
