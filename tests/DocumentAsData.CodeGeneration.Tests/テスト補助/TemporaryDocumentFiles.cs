namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// コード生成の書き込みテストで作成した一時Word文書を管理します。
/// </summary>
sealed class TemporaryDocumentFiles : IDisposable
{
    /// <summary>
    /// テスト終了時に削除する一時Word文書のパスです。
    /// </summary>
    readonly List<string> filePaths = [];

    /// <summary>
    /// テストで保存先として使用する一意なWord文書パスを作成します。
    /// </summary>
    /// <returns>まだファイルが存在しない一時Word文書のパス。</returns>
    internal string NewFilePath()
    {
        var directory = Path.Combine(
            Path.GetTempPath(),
            "DocumentAsData.CodeGeneration.Tests");
        Directory.CreateDirectory(directory);

        var filePath = Path.Combine(directory, $"{Guid.NewGuid():N}.docx");

        filePaths.Add(filePath);

        return filePath;
    }

    /// <summary>
    /// このテストで作成された一時Word文書を削除します。
    /// </summary>
    public void Dispose()
    {
        foreach (var filePath in filePaths.Where(File.Exists))
        {
            File.Delete(filePath);
        }
    }
}
