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

        var typeName = $"{Path.GetFileNameWithoutExtension(filePath)}Document";
        using var document = Document.Open(filePath);
        var mergeFieldProperties = string.Concat(
            from mergeField in document.MergeFields
            select MergeFieldPropertyDeclaration(mergeField));

        return
        [
            $$"""
            using Marimo.DocumentAsData;

            namespace {{options.Namespace}};

            public partial class {{typeName}} : Document
            {
                {{typeName}}(string filePath)
                    : base(filePath)
                {
                }

                public static new {{typeName}} Open(string filePath) =>
                    new(filePath);
            {{mergeFieldProperties}}
            }
            """
        ];
    }

    /// <summary>
    /// 生成Document型から指定したMERGEFIELDを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="mergeField">プロパティとして公開するMERGEFIELD。</param>
    /// <returns>MERGEFIELDプロパティのC#コード。</returns>
    static string MergeFieldPropertyDeclaration(MergeField mergeField) =>
        $$"""

            public string {{mergeField.Name}} =>
                MergeFields["{{mergeField.Name}}"].Text;
        """;

    /// <summary>
    /// 指定したWord文書を解析し、コード生成前に検出できる問題を診断します。
    /// </summary>
    /// <param name="filePath">診断対象のWord文書のパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>検出された診断情報。</returns>
    public static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        [];
}
