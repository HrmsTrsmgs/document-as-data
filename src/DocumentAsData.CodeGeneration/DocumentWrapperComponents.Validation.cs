namespace Marimo.DocumentAsData.CodeGeneration;

static partial class DocumentWrapperComponents
{
    /// <summary>
    /// 型付き文書を開くときの必須項目検査を生成します。
    /// 文書全体の検査と明細ごとの検査をまとめ、宣言の生成とは分離します。
    /// </summary>
    static class Validation
    {
        /// <summary>
        /// 文書直下のMERGEFIELDと文字列Content Controlから名前を取得する生成コードです。
        /// </summary>
        const string TopLevelTextItemNames =
            """
            [
                .. from item in MergeFields
                   where !item.IsInRepeatingSection
                   select item.Name,
                .. from item in ContentControls
                   where !item.IsInRepeatingSection
                   select item.Tag
            ]
            """;

        /// <summary>
        /// 指定したContent Controlコレクションから文書直下のTagを取得する生成コードです。
        /// </summary>
        /// <param name="collectionName">生成コード内のContent Controlコレクション名。</param>
        /// <returns>文書直下のTagを取得する生成コード。</returns>
        static string TopLevelTagNames(string collectionName) =>
            $$"""
            [
                .. from item in {{collectionName}}
                   where !item.IsInRepeatingSection
                   select item.Tag
            ]
            """;

        /// <summary>
        /// 繰り返し明細内の指定したContent ControlからTagを取得する生成コードです。
        /// </summary>
        /// <param name="collectionName">生成コード内のContent Controlコレクション名。</param>
        /// <returns>明細内のTagを取得する生成コード。</returns>
        static string RepeatingItemTagNames(string collectionName) =>
            $$"""
            [
                .. from control in item.{{collectionName}}
                   select control.Tag
            ]
            """;

        /// <summary>
        /// 繰り返し明細内のMERGEFIELDと文字列Content Controlから名前を取得する生成コードです。
        /// </summary>
        const string RepeatingItemTextItemNames =
            """
            [
                .. from field in item.MergeFields
                   select field.Name,
                .. from control in item.ContentControls
                   select control.Tag
            ]
            """;

        /// <summary>
        /// 生成元の文書に必要な項目を、種類ごと、次に明細ごとの順で検査するメソッドを生成します。
        /// </summary>
        /// <param name="document">必要な項目の見本を持つ生成元の文書。</param>
        /// <returns>型付き文書内に配置する検査メソッドのC#コード。</returns>
        internal static string Declaration(Document document) =>
            $$"""
                /// <summary>
                /// 生成元の文書にあったMERGEFIELD、文字列Content Control、CheckBox、DatePickerが存在することを検証します。
                /// 文字列項目は同名のMERGEFIELDと文字列Content Controlを読み取り先として認めます。
                /// 繰り返しセクション自体と、その中の文字列Content Controlも検証します。
                /// 文書の生成、返却、解放は呼び出し元のOpenで行います。
                /// </summary>
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// </exception>
                void ValidateRequiredItems()
                {
            {{RequiredItemsValidation(
                [.. from mergeField in DocumentItemScopes.TopLevel(document.MergeFields) select mergeField.Name],
                TopLevelTextItemNames,
                "MERGEFIELD")}}
            {{RequiredItemsValidation(
                [.. from contentControl in DocumentItemScopes.TopLevel(document.ContentControls) select contentControl.Tag],
                TopLevelTextItemNames,
                "文字列Content Control")}}
            {{RequiredItemsValidation(
                [.. from checkBox in DocumentItemScopes.TopLevel(document.CheckBoxes) select checkBox.Tag],
                TopLevelTagNames("CheckBoxes"),
                "CheckBox")}}
            {{RequiredItemsValidation(
                [.. from datePicker in DocumentItemScopes.TopLevel(document.DatePickers) select datePicker.Tag],
                TopLevelTagNames("DatePickers"),
                "DatePicker")}}
            {{RequiredItemsValidation(
                [.. from section in document.RepeatingSections select section.Tag],
                "RepeatingSections.Select(it => it.Tag)",
                "繰り返しセクション")}}
            {{ForEach(
                from section in document.RepeatingSections
                select RepeatingSectionValidation(section))}}
                }
            """;

        /// <summary>
        /// 各明細の文字列Content Controlを、生成元の最初の明細と照合する検査を生成します。
        /// </summary>
        /// <param name="section">必要な項目の見本を持つ生成元のセクション。</param>
        /// <returns>各明細に既存の必須項目検査を適用するC#コード。</returns>
        static string RepeatingSectionValidation(RepeatingSection section) =>
            $$"""
                    foreach (var item in RepeatingSections[{{StringLiteral(section.Tag)}}].Items)
                    {
            {{Indent(RequiredItemsValidation(
                [.. from contentControl in section.Items[0].ContentControls select contentControl.Tag],
                RepeatingItemTextItemNames,
                "文字列Content Control"))}}
            {{Indent(RequiredItemsValidation(
                [.. from mergeField in section.Items[0].MergeFields select mergeField.Name],
                RepeatingItemTextItemNames,
                "MERGEFIELD"))}}
            {{Indent(RequiredItemsValidation(
                [.. from checkBox in section.Items[0].CheckBoxes select checkBox.Tag],
                RepeatingItemTagNames("CheckBoxes"),
                "CheckBox"))}}
            {{Indent(RequiredItemsValidation(
                [.. from datePicker in section.Items[0].DatePickers select datePicker.Tag],
                RepeatingItemTagNames("DatePickers"),
                "DatePicker"))}}
                    }
            """;

        /// <summary>
        /// 不足した名前と種類を例外メッセージへ含める検査コードを生成します。
        /// </summary>
        /// <param name="requiredNames">生成元の文書に存在する、その種類の項目名またはTag。</param>
        /// <param name="actualNamesExpression">開いた文書から照合先の名前を列挙するC#式。</param>
        /// <param name="itemKind">メッセージに表示する項目の種類。</param>
        /// <returns>詳細メッセージ付きの検査コード。対象が0件なら空文字列。</returns>
        static string RequiredItemsValidation(
            string[] requiredNames,
            string actualNamesExpression,
            string itemKind) =>
            requiredNames.Length == 0
                ? ""
                : $$"""
                          {
                              var missingNames = new[] { {{string.Join(", ", requiredNames.Select(StringLiteral))}} }
                                  .Except({{actualNamesExpression}})
                                  .ToArray();

                              if (missingNames.Length > 0)
                              {
                                  throw new DocumentMappingException(
                                      {{StringLiteral($"{itemKind}が不足しています: ")}} + string.Join(", ", missingNames));
                              }
                          }
                  """;

        /// <summary>
        /// 文書直下と同じ検査ブロックを明細ループ内へ配置できるよう、各行を1段深くします。
        /// </summary>
        /// <param name="source">インデントを追加する生成コード。</param>
        /// <returns>各行に4文字の空白を追加したC#コード。</returns>
        static string Indent(string source) =>
            "    " + source.Replace(Environment.NewLine, Environment.NewLine + "    ");
    }
}
