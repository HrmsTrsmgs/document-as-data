namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の構造化されたデータ項目を表します。
/// </summary>
public abstract class DocumentItem
{
    /// <summary>
    /// 指定した文書に属するデータ項目を作成します。
    /// </summary>
    /// <param name="document">データ項目が属する文書。</param>
    private protected DocumentItem(Document document)
    {
        Document = document;
    }

    /// <summary>
    /// データ項目が属する文書を取得します。
    /// </summary>
    public Document Document { get; }
}
