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
    /// 生成したC#ソースファイルを取得します。
    /// </summary>
    [Output]
    public ITaskItem[] GeneratedFiles { get; set; } = [];

    /// <summary>
    /// コード生成を実行し、報告済みの辞書JSONエラーはタスクの失敗として返します。
    /// </summary>
    /// <returns>コード生成に成功した場合は <see langword="true"/>。</returns>
    public override bool Execute()
    {
        try
        {
            GeneratedFiles =
            [
                .. from documentFile in DocumentFiles
                   select Generate(documentFile)
            ];

            return true;
        }
        catch (JsonException) when (Log.HasLoggedErrors)
        {
            return false;
        }
    }

    /// <summary>
    /// 一つのWord文書からC#ソースを生成します。
    /// </summary>
    /// <param name="documentFile">コード生成対象のWord文書。</param>
    /// <returns>生成したC#ソースファイルを表すMSBuild項目。</returns>
    ITaskItem Generate(ITaskItem documentFile)
    {
        var documentFilePath = Path.GetFullPath(documentFile.ItemSpec);
        var generatedFilePath = Path.Combine(
            Path.GetDirectoryName(documentFilePath)!,
            $"{Path.GetFileNameWithoutExtension(documentFilePath)}.DocumentAsData.g.cs");
        var nameMappings = LoadNameMappings(documentFilePath);

        File.WriteAllText(
            generatedFilePath,
            DocumentWrapperGenerator
                .GenerateSources(
                    documentFilePath,
                    options => options.NameMappings = nameMappings)
                .Single());

        var generatedFile = new TaskItem(generatedFilePath);
        generatedFile.SetMetadata(
            "DependentUpon",
            Path.GetFileName(documentFilePath));
        generatedFile.SetMetadata("DesignTimeSharedInput", "true");
        return generatedFile;
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
