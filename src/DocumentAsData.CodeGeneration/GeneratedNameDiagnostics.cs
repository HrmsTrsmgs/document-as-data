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
        var documentProperties = DocumentPropertySourceNames(document);
        (string Tag, string[] PropertyNames)[] sections =
        [
            .. from section in document.RepeatingSections
               select (section.Tag, (from itemGroup in DetailItems(section)
                                     from name in itemGroup.Names
                                     select name).ToArray())
        ];

        var documentName = Path.GetFileNameWithoutExtension(filePath);
        HashSet<string> reservedNames =
        [
            with(StringComparer.Ordinal),
            .. ReservedMemberNames,
            options.DocumentTypeName(documentName).IdentifierComparisonKey,
            options.DataTypeName(documentName).IdentifierComparisonKey
        ];

        return
        [
            .. InvalidDocumentNameDiagnostics(documentName, options),
            .. PropertyNameDiagnostics(documentProperties, reservedNames, options, reserveRead: true),
            .. from section in sections
               from diagnostic in
                   PropertyNameDiagnostics(section.PropertyNames, reservedNames, options, reserveRead: false)
               select diagnostic,
            .. from section in sections
               let generatedTypeName = options.DataTypeName(section.Tag).IdentifierComparisonKey
               from propertySection in sections
               where options.GeneratedName(propertySection.Tag).IdentifierComparisonKey == generatedTypeName
               select new CodeGenerationDiagnostic(
                   true,
                   generatedTypeName,
                   [section.Tag, propertySection.Tag]),
            .. from section in sections
               let generatedTypeName = options.DataTypeName(section.Tag).IdentifierComparisonKey
               from sourceNames in
                   from sourceName in section.PropertyNames
                   group sourceName by options.GeneratedName(sourceName).IdentifierComparisonKey
               where sourceNames.Key == generatedTypeName
               select new CodeGenerationDiagnostic(
                   true,
                   generatedTypeName,
                   [.. sourceNames]),
            .. from section in sections
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
    /// 文書型のプロパティとして生成する通常項目名とセクション名を宣言順に取得します。
    /// </summary>
    static string[] DocumentPropertySourceNames(Document document) =>
        [
            .. from itemGroup in TopLevelItems(document)
               from name in itemGroup.Names
               select name,
            .. from section in document.RepeatingSections
               select section.Tag
        ];

    /// <summary>
    /// 同じ生成先に属する元名だけをまとめ、重複・空名・予約名との衝突を診断します。
    /// </summary>
    /// <param name="propertyNames">同じ型へ生成するプロパティの元名。</param>
    /// <param name="reservedNames">既存メンバー名と生成型名。</param>
    /// <param name="options">識別子への変換設定。</param>
    /// <param name="reserveRead">文書型ではReadを予約し、明細型ではプロパティ名として許します。</param>
    /// <returns>元名の順に並ぶ診断。</returns>
    static IEnumerable<CodeGenerationDiagnostic> PropertyNameDiagnostics(
        IEnumerable<string> propertyNames,
        HashSet<string> reservedNames,
        CodeGenerationOptions options,
        bool reserveRead) =>
        from sourceName in propertyNames
        group sourceName by options.GeneratedName(sourceName).IdentifierComparisonKey
        into sourceNames
        where sourceNames.Count() > 1
            || sourceNames.Key.IsEmpty()
            || reservedNames.Contains(sourceNames.Key)
                && (reserveRead || sourceNames.Key != nameof(Document.Read))
        select new CodeGenerationDiagnostic(
            true,
            sourceNames.Key,
            [.. sourceNames]);

    /// <summary>
    /// 生成Document型の名前を作れない文書ファイル名を診断します。
    /// </summary>
    /// <param name="sourceName">拡張子を除いた生成元の文書名。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>文書名が無効な場合は、その名前を示す診断。</returns>
    static IEnumerable<CodeGenerationDiagnostic> InvalidDocumentNameDiagnostics(
        string sourceName,
        CodeGenerationOptions options)
    {
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
