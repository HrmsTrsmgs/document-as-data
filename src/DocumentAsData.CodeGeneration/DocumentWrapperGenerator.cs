namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Word文書から、DocumentAsDataの型付きラッパーコードを生成します。
/// </summary>
public static class DocumentWrapperGenerator
{
    /// <summary>
    /// 既存の操作APIや生成メソッドとの重複、生成コードが参照するコレクションの隠蔽を防ぐために予約する名前です。
    /// </summary>
    static readonly string[] ReservedMemberNames =
    [
        nameof(Document.Read),
        nameof(Document.Open),
        nameof(Document.Replace),
        nameof(Document.Save),
        nameof(Document.SaveAs),
        nameof(Document.Close),
        nameof(Document.Dispose),
        nameof(Document.MergeFields),
        nameof(Document.ContentControls),
        nameof(Document.CheckBoxes),
        nameof(Document.DatePickers),
        nameof(Document.RepeatingSections),
        "ValidateRequiredItems"
    ];

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
        IEnumerable<string>[] propertySourceNameScopes =
        [
            [
                .. from mergeField in document.MergeFields
                   where !mergeField.IsInRepeatingSection
                   select mergeField.Name,
                .. from contentControl in document.ContentControls
                   where !contentControl.IsInRepeatingSection
                   select contentControl.Tag,
                .. from checkBox in document.CheckBoxes
                   where !checkBox.IsInRepeatingSection
                   select checkBox.Tag,
                .. from datePicker in document.DatePickers
                   where !datePicker.IsInRepeatingSection
                   select datePicker.Tag,
                .. from section in document.RepeatingSections
                   select section.Tag
            ],
            .. from section in document.RepeatingSections
               select from contentControl in section.Items[0].ContentControls
                      select contentControl.Tag
        ];
        var documentName = Path.GetFileNameWithoutExtension(filePath);
        string[] reservedNames =
        [
            .. ReservedMemberNames,
            options.DocumentTypeName(documentName).IdentifierComparisonKey,
            options.DataTypeName(documentName).IdentifierComparisonKey
        ];

        return
        [
            .. InvalidDocumentNameDiagnostics(filePath, options),
            .. from scope in propertySourceNameScopes
               from sourceNames in
                   (from sourceName in scope
                    group sourceName by options.GeneratedName(sourceName).IdentifierComparisonKey)
               where sourceNames.Count() > 1
                   || sourceNames.Key.IsEmpty()
                   || reservedNames.Contains(sourceNames.Key)
               select new CodeGenerationDiagnostic(
                   true,
                   sourceNames.Key,
                   [.. sourceNames]),
            .. from section in document.RepeatingSections
               group section.Tag by options.DataTypeName(section.Tag).IdentifierComparisonKey
               into sourceNames
               where sourceNames.Count() > 1
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
