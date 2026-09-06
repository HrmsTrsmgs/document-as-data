using System.Collections;
using Marimo.DocumentAsData.Build;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// MSBuild連携のテストで、利用者プロジェクトに近い一時ディレクトリを扱います。
/// </summary>
sealed class MSBuild連携テストプロジェクト : IDisposable
{
    const string BasicStructureDocumentFilePath =
        @"TestData\コード生成\BasicStructure.docx";

    /// <summary>
    /// テストごとに分離した一時プロジェクトを作成します。
    /// </summary>
    MSBuild連携テストプロジェクト()
    {
        DirectoryPath = Path.Combine(
            Path.GetTempPath(),
            "DocumentAsData.Tests",
            Guid.NewGuid().ToString("N"));
    }

    /// <summary>
    /// テスト対象プロジェクトのルートディレクトリです。
    /// </summary>
    internal string DirectoryPath { get; }

    /// <summary>
    /// 新しい一時プロジェクトを作成します。
    /// </summary>
    /// <returns>作成した一時プロジェクト。</returns>
    internal static MSBuild連携テストプロジェクト Create() =>
        new();

    /// <summary>
    /// テスト中に作成した一時プロジェクトを削除します。
    /// </summary>
    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
        {
            Directory.Delete(DirectoryPath, true);
        }
    }

    /// <summary>
    /// 既存の基本構造テスト文書を、指定した相対パスへ追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>追加したWord文書の絶対パス。</returns>
    internal string AddBasicStructureDocument(string relativePath)
    {
        var destinationPath = Path.Combine(DirectoryPath, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        File.Copy(BasicStructureDocumentFilePath, destinationPath);
        return destinationPath;
    }

    /// <summary>
    /// MSBuildタスクを直接実行します。
    /// </summary>
    /// <param name="documentFilePaths">生成対象Word文書の絶対パス。</param>
    /// <returns>タスクの実行結果。</returns>
    internal MSBuild連携タスク実行結果 Generate(params string[] documentFilePaths)
    {
        var buildEngine = new RecordingBuildEngine();
        var task = new GenerateDocumentAsData
        {
            BuildEngine = buildEngine,
            DocumentFiles =
            [
                .. from filePath in documentFilePaths
                   select new TaskItem(filePath)
            ]
        };

        return new(
            task.Execute(),
            task.GeneratedFiles,
            buildEngine.Warnings);
    }

    /// <summary>
    /// 指定したWord文書に対応する生成ファイルの絶対パスを取得します。
    /// </summary>
    /// <param name="documentRelativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>生成ファイルの絶対パス。</returns>
    internal string GeneratedFilePathFor(string documentRelativePath) =>
        Path.Combine(
            DirectoryPath,
            Path.GetDirectoryName(documentRelativePath) ?? "",
            $"{Path.GetFileNameWithoutExtension(documentRelativePath)}.DocumentAsData.g.cs");
}

/// <summary>
/// MSBuildタスクの実行結果を、テスト内で観測しやすい形にまとめます。
/// </summary>
/// <param name="Succeeded">タスクが成功した場合は <see langword="true"/>。</param>
/// <param name="GeneratedFiles">タスクが返した生成ファイル。</param>
/// <param name="Warnings">タスクが記録した警告。</param>
sealed record MSBuild連携タスク実行結果(
    bool Succeeded,
    ITaskItem[] GeneratedFiles,
    IReadOnlyList<BuildWarningEventArgs> Warnings)
{
    /// <summary>
    /// 生成ファイルの絶対パス一覧です。
    /// </summary>
    internal string[] GeneratedFilePaths =>
        [.. GeneratedFiles.Select(it => it.ItemSpec)];

    /// <summary>
    /// 生成ファイルが一つであるテストで、そのファイルパスを取得します。
    /// </summary>
    internal string SingleGeneratedFilePath =>
        GeneratedFilePaths.Single();

    /// <summary>
    /// 生成ファイルが一つであるテストで、そのC#ソースを取得します。
    /// </summary>
    internal string SingleGeneratedSource =>
        File.ReadAllText(SingleGeneratedFilePath);
}

/// <summary>
/// MSBuildタスクが出力した警告を記録します。
/// </summary>
sealed class RecordingBuildEngine : IBuildEngine
{
    /// <summary>
    /// タスクが記録した警告です。
    /// </summary>
    internal List<BuildWarningEventArgs> Warnings { get; } = [];

    /// <summary>
    /// タスクエラー後も処理を継続するかどうかを取得します。
    /// </summary>
    public bool ContinueOnError => false;

    /// <summary>
    /// テスト用のタスク要素行番号を取得します。
    /// </summary>
    public int LineNumberOfTaskNode => 0;

    /// <summary>
    /// テスト用のタスク要素列番号を取得します。
    /// </summary>
    public int ColumnNumberOfTaskNode => 0;

    /// <summary>
    /// テスト用のプロジェクトファイル名を取得します。
    /// </summary>
    public string ProjectFileOfTaskNode => "";

    /// <summary>
    /// このテスト補助では別のMSBuildプロジェクトをビルドしません。
    /// </summary>
    public bool BuildProjectFile(
        string projectFileName,
        string[] targetNames,
        IDictionary globalProperties,
        IDictionary targetOutputs) =>
        throw new NotImplementedException();

    /// <summary>
    /// カスタムイベントを受け取ります。
    /// </summary>
    public void LogCustomEvent(CustomBuildEventArgs e)
    {
    }

    /// <summary>
    /// エラーイベントを受け取ります。
    /// </summary>
    public void LogErrorEvent(BuildErrorEventArgs e)
    {
    }

    /// <summary>
    /// メッセージイベントを受け取ります。
    /// </summary>
    public void LogMessageEvent(BuildMessageEventArgs e)
    {
    }

    /// <summary>
    /// 警告イベントを記録します。
    /// </summary>
    public void LogWarningEvent(BuildWarningEventArgs e)
    {
        Warnings.Add(e);
    }
}
