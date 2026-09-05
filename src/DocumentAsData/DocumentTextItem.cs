namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内で文字列の値を読み書きできるデータ項目を表します。
/// </summary>
public abstract class DocumentTextItem
{
    /// <summary>
    /// 指定した文書に属するデータ項目を作成します。
    /// </summary>
    /// <param name="document">データ項目が属する文書。</param>
    protected DocumentTextItem(Document document)
    {
        Document = document;
    }

    /// <summary>
    /// データ項目が属する文書を取得します。
    /// </summary>
    public Document Document { get; }

    /// <summary>
    /// データ項目の文字列値を取得または設定します。
    /// </summary>
    public abstract string Value { get; set; }
}
