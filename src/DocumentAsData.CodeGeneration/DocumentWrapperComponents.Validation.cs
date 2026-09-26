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
        /// 生成元の文書に必要な項目を、種類ごと、次に明細ごとの順で検査するメソッドを生成します。
        /// </summary>
        /// <param name="document">必要な項目の見本を持つ生成元の文書。</param>
        /// <returns>型付き文書内に配置する検査メソッドのC#コード。</returns>
        internal static string Declaration(Document document) =>
            $$"""
                /// <summary>
                /// 生成元の文書にあったMERGEFIELD、文字列Content Control、CheckBox、DatePickerが存在することを検証します。
                /// 文字列項目は同名のMERGEFIELDと文字列Content Controlを読み取り先として認めます。
                /// 繰り返しセクション内の文字列Content Controlは、明細ごとにも存在を検証します。
                /// 文書の生成、返却、解放は呼び出し元のOpenで行います。
                /// </summary>
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// </exception>
                void ValidateRequiredItems()
                {
            {{RequiredItemsValidation(
                [.. from mergeField in document.MergeFields select mergeField.Name],
                "[.. MergeFields.Select(it => it.Name), .. ContentControls.Select(it => it.Tag)]",
                "MERGEFIELD")}}
            {{RequiredItemsValidation(
                [.. from contentControl in document.ContentControls select contentControl.Tag],
                "[.. ContentControls.Select(it => it.Tag), .. MergeFields.Select(it => it.Name)]",
                "文字列Content Control")}}
            {{RequiredItemsValidation(
                [.. from checkBox in document.CheckBoxes select checkBox.Tag],
                "CheckBoxes.Select(it => it.Tag)",
                "CheckBox")}}
            {{RequiredItemsValidation(
                [.. from datePicker in document.DatePickers select datePicker.Tag],
                "DatePickers.Select(it => it.Tag)",
                "DatePicker")}}
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
                "item.ContentControls.Select(it => it.Tag)",
                "文字列Content Control"))}}
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
