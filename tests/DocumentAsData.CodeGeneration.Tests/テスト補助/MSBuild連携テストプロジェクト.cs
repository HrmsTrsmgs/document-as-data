using System.Collections;
using System.Diagnostics;
using System.Security;
using System.Text;
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
    const string CustomerDataDocumentFilePath =
        @"TestData\コード生成\customerData.docx";

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
    /// リポジトリのtargetsとビルド済みタスクを使う一時プロジェクトを作ります。パッケージの配置は対象外です。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellGenerationSample()
    {
        var targetsPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "src", "DocumentAsData.Build", "buildTransitive",
            "Marimo.DocumentAsData.Build.targets"));
        File.WriteAllText(
            Path.Combine(DirectoryPath, "DocumentAsData.Generate.proj"),
            $$"""
            <Project>
              <PropertyGroup>
                <DocumentAsDataTaskAssembly>{{SecurityElement.Escape(typeof(GenerateDocumentAsData).Assembly.Location)}}</DocumentAsDataTaskAssembly>
              </PropertyGroup>
              <ItemGroup>
                <DocumentAsData Include="BasicStructure.docx" />
              </ItemGroup>
              <Import Project="{{SecurityElement.Escape(targetsPath)}}" />
              <Target Name="Build" DependsOnTargets="GenerateDocumentAsDataSources" />
            </Project>
            """);
        var scriptFilePath = Path.Combine(DirectoryPath, "Generate.ps1");
        File.WriteAllText(
            scriptFilePath,
            """
            $ErrorActionPreference = 'Stop'
            [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            dotnet msbuild ./DocumentAsData.Generate.proj /t:Build /nologo /v:minimal
            exit $LASTEXITCODE
            """);
        return scriptFilePath;
    }

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
    internal string AddBasicStructureDocument(string relativePath) =>
        AddDocument(BasicStructureDocumentFilePath, relativePath);

    /// <summary>
    /// 既存の顧客データテスト文書を、指定した相対パスへ追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>追加したWord文書の絶対パス。</returns>
    internal string AddCustomerDataDocument(string relativePath) =>
        AddDocument(CustomerDataDocumentFilePath, relativePath);

    /// <summary>
    /// プロジェクト直下へ、指定したWord文書用の識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="documentFileName">対応するWord文書のファイル名。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddProjectDictionaryFor(
        string documentFileName,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                DictionaryFileName(documentFileName)),
            json);

    /// <summary>
    /// Word文書と同じディレクトリへ、識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="documentRelativePath">対応するWord文書の一時プロジェクト内相対パス。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddDocumentDictionaryFor(
        string documentRelativePath,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                Path.GetDirectoryName(documentRelativePath) ?? "",
                DictionaryFileName(documentRelativePath)),
            json);

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
            ],
            ProjectDirectory = DirectoryPath
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

    /// <summary>
    /// 既存のWord文書を一時プロジェクトへコピーします。
    /// </summary>
    /// <param name="sourcePath">コピー元のWord文書。</param>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>コピーしたWord文書の絶対パス。</returns>
    string AddDocument(
        string sourcePath,
        string relativePath)
    {
        var destinationPath = Path.Combine(DirectoryPath, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        File.Copy(sourcePath, destinationPath);
        return destinationPath;
    }

    /// <summary>
    /// 識別子名変換辞書を書き込みます。
    /// </summary>
    /// <param name="filePath">書き込み先ファイル。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    static string WriteDictionary(
        string filePath,
        string json)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
        File.WriteAllText(filePath, json);
        return filePath;
    }

    /// <summary>
    /// Word文書名に対応する識別子名変換辞書のファイル名を作ります。
    /// </summary>
    /// <param name="documentFileName">対応するWord文書のファイル名。</param>
    /// <returns>識別子名変換辞書のファイル名。</returns>
    static string DictionaryFileName(string documentFileName) =>
        $"{Path.GetFileNameWithoutExtension(documentFileName)}.documentasdata.json";
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
    /// 生成ファイルが一つであるテストで、そのMSBuild項目を取得します。
    /// </summary>
    internal ITaskItem SingleGeneratedFile =>
        GeneratedFiles.Single();

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
/// PowerShellの終了コードと診断出力を保持します。
/// </summary>
/// <param name="ExitCode">プロセスの終了コード。</param>
/// <param name="Output">標準出力と標準エラー。</param>
sealed record PowerShell実行結果(int ExitCode, string Output)
{
    /// <summary>
    /// スクリプトを別プロセスで実行し、両方の出力を同時に読み出してバッファ待ちを防ぎます。
    /// </summary>
    /// <param name="scriptFilePath">実行するスクリプト。</param>
    /// <param name="workingDirectory">MSBuildの作業ディレクトリ。</param>
    /// <returns>終了コードと出力。</returns>
    internal static PowerShell実行結果 Run(string scriptFilePath, string workingDirectory)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "pwsh",
                ArgumentList = { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFilePath },
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };
        process.Start();
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        process.WaitForExit();
        return new(process.ExitCode, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
    }
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
