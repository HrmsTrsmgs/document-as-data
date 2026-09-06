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
    /// 生成したC#ソースファイルを取得します。
    /// </summary>
    [Output]
    public ITaskItem[] GeneratedFiles { get; set; } = [];

    /// <summary>
    /// コード生成を実行します。
    /// </summary>
    /// <returns>コード生成に成功した場合は <see langword="true"/>。</returns>
    public override bool Execute()
    {
        GeneratedFiles =
        [
            .. from documentFile in DocumentFiles
               select Generate(documentFile)
        ];

        return true;
    }

    /// <summary>
    /// 一つのWord文書からC#ソースを生成します。
    /// </summary>
    /// <param name="documentFile">コード生成対象のWord文書。</param>
    /// <returns>生成したC#ソースファイルを表すMSBuild項目。</returns>
    static ITaskItem Generate(ITaskItem documentFile)
    {
        var documentFilePath = Path.GetFullPath(documentFile.ItemSpec);
        var generatedFilePath = Path.Combine(
            Path.GetDirectoryName(documentFilePath)!,
            $"{Path.GetFileNameWithoutExtension(documentFilePath)}.DocumentAsData.g.cs");

        File.WriteAllText(
            generatedFilePath,
            DocumentWrapperGenerator.GenerateSources(documentFilePath).Single());

        var generatedFile = new TaskItem(generatedFilePath);
        generatedFile.SetMetadata(
            "DependentUpon",
            Path.GetFileName(documentFilePath));
        generatedFile.SetMetadata("DesignTimeSharedInput", "true");
        return generatedFile;
    }
}
