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
        configure?.Invoke(options);

        using var document = Document.Open(filePath);
        string[] propertySourceNames =
        [
            .. from mergeField in document.MergeFields
               select mergeField.Name,
            .. from contentControl in document.ContentControls
               select contentControl.Tag,
            .. from checkBox in document.CheckBoxes
               select checkBox.Tag,
            .. from datePicker in document.DatePickers
               select datePicker.Tag
        ];

        return
        [
            .. InvalidDocumentNameDiagnostics(filePath, options),
            .. from sourceName in propertySourceNames
               group sourceName by options.GeneratedName(sourceName)
               into sourceNames
               // 生成するメソッドとの重複や、生成コードが参照するコレクションの隠蔽も診断します。
               where sourceNames.Count() > 1
                   || sourceNames.Key is nameof(Document.Read)
                       or nameof(Document.Open)
                       or nameof(Document.Replace)
                       or nameof(Document.MergeFields)
               select new CodeGenerationDiagnostic(
                   true,
                   sourceNames.Key,
                   [.. sourceNames])
        ];
    }

    /// <summary>
    /// 生成Document型の名前を作れない文書ファイル名を診断します。
    /// </summary>
    /// <param name="filePath">診断対象のWord文書のパス。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>文書名が無効な場合は、その名前を示す診断。</returns>
    static IEnumerable<CodeGenerationDiagnostic> InvalidDocumentNameDiagnostics(
        string filePath,
        CodeGenerationOptions options)
    {
        var sourceName = Path.GetFileNameWithoutExtension(filePath);
        var generatedName = options.GeneratedName(sourceName);

        if (!generatedName.IsEmpty())
        {
            yield break;
        }

        yield return new CodeGenerationDiagnostic(
            true,
            generatedName,
            [sourceName],
            sourceName);
    }
}
