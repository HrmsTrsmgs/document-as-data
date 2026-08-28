namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを表します。
/// </summary>
public class ContentControl
{
    /// <summary>
    /// Content Controlが属する文書を取得します。
    /// </summary>
    public Document Document =>
        throw new NotImplementedException();

    /// <summary>
    /// Content ControlのTagを取得します。
    /// </summary>
    public string Tag =>
        throw new NotImplementedException();

    /// <summary>
    /// Content Controlの値を取得または設定します。
    /// </summary>
    public string Value
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }
}
