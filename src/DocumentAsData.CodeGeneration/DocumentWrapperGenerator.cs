namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Word文書から、DocumentAsDataの型付きラッパーコードを生成します。
/// </summary>
public static class DocumentWrapperGenerator
{
    /// <summary>
    /// 指定したWord文書からC#ソースコードを生成します。
    /// </summary>
    /// <param name="filePath">生成元のWord文書のパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    public static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null)
    {
        var options = new CodeGenerationOptions();
        configure?.Invoke(options);

        using var document = Document.Open(filePath);
        return
        [
            DocumentWrapperComponents.SourceFile(filePath, options, document)
        ];
    }

    /// <summary>
    /// 指定したWord文書を解析し、コード生成前に検出できる問題を診断します。
    /// </summary>
    /// <param name="filePath">診断対象のWord文書のパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>検出された診断情報。</returns>
    public static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null)
    {
        var options = new CodeGenerationOptions();

        using var document = Document.Open(filePath);
        return
        [
            .. from mergeField in document.MergeFields
               group mergeField.Name by options.GeneratedName(mergeField.Name)
               into sourceNames
               where sourceNames.Count() > 1
               select new CodeGenerationDiagnostic(
                   true,
                   sourceNames.Key,
                   [.. sourceNames])
        ];
    }
}
