namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内で文字列の値を読み書きできるデータ項目を表します。
/// </summary>
interface DocumentTextItem
{
    /// <summary>
    /// データ項目の文字列を取得または設定します。
    /// </summary>
    string Text { get; set; }
}
