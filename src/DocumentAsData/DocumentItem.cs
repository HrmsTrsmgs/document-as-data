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

    /// <summary>
    /// 名前やTagを引用符で囲み、引用符、バックスラッシュ、改行（LF）をエスケープして表示します。
    /// 日本語などの通常の文字は、そのまま読める形で保持します。
    /// </summary>
    /// <param name="name">表示する名前またはTag。</param>
    /// <returns>引用符で囲んだ表示用文字列。</returns>
    private protected static string QuoteName(string name) =>
        "\"" + name
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n") + "\"";
}
