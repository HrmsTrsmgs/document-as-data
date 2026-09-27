using static Marimo.DocumentAsData.CodeGeneration.DocumentItemScopes;

namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// 生成する型とプロパティの元名を、生成先の範囲ごとに診断します。
/// 文書を開く責務やC#コードの出力とは分離します。
/// </summary>
static class GeneratedNameDiagnostics
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
    /// 開いた文書の生成名を、文書名・プロパティ・明細型名の順に診断します。
    /// </summary>
    /// <param name="filePath">生成元の文書パス。</param>
    /// <param name="document">診断対象の文書。</param>
    /// <param name="options">名前変換の設定。</param>
    /// <returns>診断結果。文書の所有権は呼び出し元に残します。</returns>
    internal static CodeGenerationDiagnostic[] Diagnose(
        string filePath,
        Document document,
        CodeGenerationOptions options)
    {
        var propertySourceNameScopes = PropertySourceNameScopes(document);
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
            .. PropertyNameDiagnostics(propertySourceNameScopes, reservedNames, options),
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
    /// 通常項目とセクション名を文書型へ、最初の明細の文字列Content Controlを各明細型へ分けます。
    /// 明細型では文字列Content Controlのみが現在の診断対象です。
    /// </summary>
    /// <param name="document">生成元の文書。</param>
    /// <returns>同じ生成先へ配置される元名の列。</returns>
    static IEnumerable<string>[] PropertySourceNameScopes(Document document) =>
        [
            [
                .. from mergeField in TopLevel(document.MergeFields)
                   select mergeField.Name,
                .. from contentControl in TopLevel(document.ContentControls)
                   select contentControl.Tag,
                .. from checkBox in TopLevel(document.CheckBoxes)
                   select checkBox.Tag,
                .. from datePicker in TopLevel(document.DatePickers)
                   select datePicker.Tag,
                .. from section in document.RepeatingSections
                   select section.Tag
            ],
            .. from section in document.RepeatingSections
               select from contentControl in section.Items[0].ContentControls
                      select contentControl.Tag
        ];

    /// <summary>
    /// 同じ生成先に属する元名だけをまとめ、重複・空名・予約名との衝突を診断します。
    /// </summary>
    /// <param name="scopes">生成先ごとに分けた元名の列。</param>
    /// <param name="reservedNames">既存メンバー名と生成型名。</param>
    /// <param name="options">識別子への変換設定。</param>
    /// <returns>スコープ順・元名順の診断。</returns>
    static IEnumerable<CodeGenerationDiagnostic> PropertyNameDiagnostics(
        IEnumerable<IEnumerable<string>> scopes,
        string[] reservedNames,
        CodeGenerationOptions options) =>
        from scope in scopes
        from sourceNames in
            (from sourceName in scope
             group sourceName by options.GeneratedName(sourceName).IdentifierComparisonKey)
        where sourceNames.Count() > 1
            || sourceNames.Key.IsEmpty()
            || reservedNames.Contains(sourceNames.Key)
        select new CodeGenerationDiagnostic(
            true,
            sourceNames.Key,
            [.. sourceNames]);

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
