namespace Marimo.DocumentAsData.Test.テスト補助;

/// <summary>
/// テストで作成した一時ファイルを使用後に削除します。
/// </summary>
sealed class TemporaryFile(string filePath) : IDisposable
{
    public void Dispose() => File.Delete(filePath);
}
