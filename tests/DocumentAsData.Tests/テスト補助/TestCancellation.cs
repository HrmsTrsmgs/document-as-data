namespace Marimo.DocumentAsData.Test.テスト補助;

/// <summary>
/// テスト実行のキャンセル通知を取得します。
/// </summary>
static class TestCancellation
{
    /// <summary>
    /// 呼び出すたびに実行中のテストのキャンセルトークンを取得します。
    /// </summary>
    public static CancellationToken Token => TestContext.Current.CancellationToken;
}
