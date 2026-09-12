using System.Text.Json;
using Marimo.DocumentAsData.CodeGeneration;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Marimo.DocumentAsData.Build;

/// <summary>
/// MSBuildのDocumentAsData項目から、型付きDocument用C#ソースを生成します。
/// </summary>
public sealed class GenerateDocumentAsData : Microsoft.Build.Utilities.Task
{
    /// <summary>
    /// コード生成対象のWord文書を取得または設定します。
    /// </summary>
    [Required]
    public ITaskItem[] DocumentFiles { get; set; } = [];

    /// <summary>
    /// 対象プロジェクトのディレクトリを取得または設定します。
    /// </summary>
    [Required]
    public string ProjectDirectory { get; set; } = "";

    /// <summary>
    /// 生成コードの名前空間を取得または設定します。既定値はGeneratedです。
    /// </summary>
    public string RootNamespace { get; set; } = "Generated";

    /// <summary>
    /// 生成したC#ソースファイルを取得します。
    /// </summary>
    [Output]
    public ITaskItem[] GeneratedFiles { get; set; } = [];

    /// <summary>
    /// コード生成を実行し、診断エラー、ファイルの不在、報告済みの辞書JSONエラーはタスクの失敗として返します。
    /// </summary>
    /// <returns>コード生成に成功した場合は <see langword="true"/>。</returns>
    public override bool Execute()
    {
        try
        {
            GeneratedFiles =
            [
                .. from documentFile in DocumentFiles
                   from generatedFile in Generate(documentFile)
                   select generatedFile
            ];

            return !Log.HasLoggedErrors;
        }
        catch (JsonException) when (Log.HasLoggedErrors)
        {
            return false;
        }
        catch (FileNotFoundException exception)
        {
            Log.LogError(
                subcategory: null,
                errorCode: null,
                helpKeyword: null,
                file: exception.FileName,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: exception.Message);
            return false;
        }
    }

    /// <summary>
    /// 一つのWord文書を診断し、エラーがなければC#ソースを生成します。
    /// </summary>
    /// <param name="documentFile">コード生成対象のWord文書。</param>
    /// <returns>生成したC#ソースファイルを表すMSBuild項目。診断エラーがある場合は空です。</returns>
    IEnumerable<ITaskItem> Generate(ITaskItem documentFile)
    {
        var documentFilePath = Path.GetFullPath(documentFile.ItemSpec);
        var generatedFilePath = Path.Combine(
            Path.GetDirectoryName(documentFilePath)!,
            $"{Path.GetFileNameWithoutExtension(documentFilePath)}.DocumentAsData.g.cs");
        var nameMappings = LoadNameMappings(documentFilePath);

        foreach (var diagnostic in
            from diagnostic in
                DocumentWrapperGenerator.GenerateDiagnostics(
                    documentFilePath,
                    options => ConfigureOptions(options, nameMappings))
            where diagnostic.IsError
            select diagnostic)
        {
            Log.LogError(
                subcategory: null,
                errorCode: null,
                helpKeyword: null,
                file: documentFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: $"DocumentAsData のコード生成診断: 生成名 '{diagnostic.GeneratedName}'、元名 '{string.Join(", ", diagnostic.SourceNames)}'");
        }

        if (Log.HasLoggedErrors)
        {
            yield break;
        }

        WriteIfChanged(
            generatedFilePath,
            DocumentWrapperGenerator
                .GenerateSources(
                    documentFilePath,
                    options => ConfigureOptions(options, nameMappings))
                .Single());

        var generatedFile = new TaskItem(generatedFilePath);
        generatedFile.SetMetadata(
            "DependentUpon",
            Path.GetFileName(documentFilePath));
        generatedFile.SetMetadata("DesignTimeSharedInput", "true");
        yield return generatedFile;
    }

    /// <summary>
    /// 内容が変わった場合だけ書き込み、同じ内容での更新日時の変更を防ぎます。
    /// </summary>
    /// <param name="path">生成ファイルのパス。</param>
    /// <param name="content">生成したソースコード。</param>
    static void WriteIfChanged(string path, string content)
    {
        if (File.Exists(path) && File.ReadAllText(path) == content)
        {
            return;
        }

        File.WriteAllText(path, content);
    }

    /// <summary>
    /// 診断とソース生成へ同じ名前変換辞書と名前空間を適用します。
    /// </summary>
    /// <param name="options">適用先のコード生成設定。</param>
    /// <param name="nameMappings">文書に対応する識別子名変換辞書。</param>
    void ConfigureOptions(CodeGenerationOptions options, Dictionary<string, string> nameMappings)
    {
        options.NameMappings = nameMappings;
        options.Namespace = RootNamespace;
    }

    /// <summary>
    /// 文書隣またはプロジェクト直下の識別子名変換辞書を読み込みます。
    /// 不正なJSONは辞書のパス付きで報告し、再スローして生成を中断します。
    /// </summary>
    /// <param name="documentFilePath">コード生成対象のWord文書。</param>
    /// <returns>文書内の名前から生成後のC#名への対応表。</returns>
    Dictionary<string, string> LoadNameMappings(string documentFilePath)
    {
        var dictionaryFilePath = DictionaryFilePath(documentFilePath);

        if (dictionaryFilePath is null)
        {
            return [];
        }

        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, string>>(
                File.ReadAllText(dictionaryFilePath))
                ?? [];
        }
        catch (JsonException exception)
        {
            Log.LogError(
                subcategory: null,
                errorCode: null,
                helpKeyword: null,
                file: dictionaryFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: exception.Message);
            throw;
        }
    }

    /// <summary>
    /// 文書隣を優先して、指定したWord文書用の識別子名変換辞書を探します。
    /// </summary>
    /// <param name="documentFilePath">コード生成対象のWord文書。</param>
    /// <returns>見つかった辞書ファイルのパス。存在しない場合は <see langword="null"/>。</returns>
    string? DictionaryFilePath(string documentFilePath)
    {
        var dictionaryFileName =
            $"{Path.GetFileNameWithoutExtension(documentFilePath)}.documentasdata.json";
        var sameDirectoryDictionaryPath = Path.Combine(
            Path.GetDirectoryName(documentFilePath)!,
            dictionaryFileName);

        if (File.Exists(sameDirectoryDictionaryPath))
        {
            return sameDirectoryDictionaryPath;
        }

        var projectDirectoryDictionaryPath = Path.Combine(
            ProjectDirectory,
            dictionaryFileName);

        return File.Exists(projectDirectoryDictionaryPath)
            ? projectDirectoryDictionaryPath
            : null;
    }
}
