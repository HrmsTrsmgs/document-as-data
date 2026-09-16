namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// コード生成設定から生成後の名前を決定する処理を提供します。
/// </summary>
static class CodeGenerationOptionNames
{
    extension(CodeGenerationOptions self)
    {
        /// <summary>
        /// 元文書名の名前変換へDocument接尾辞を付け、宣言と衝突診断で同じ型名を使います。
        /// </summary>
        /// <param name="documentName">拡張子を除いた元文書名。</param>
        /// <returns>型付き文書の生成型名。比較用の正規化は行いません。</returns>
        internal string DocumentTypeName(string documentName) =>
            $"{self.GeneratedName(documentName)}Document";

        /// <summary>
        /// 元文書名の名前変換へData接尾辞を付け、宣言・参照と衝突診断で同じ型名を使います。
        /// </summary>
        /// <param name="documentName">拡張子を除いた元文書名。</param>
        /// <returns>文書データの生成型名。比較用の正規化は行いません。</returns>
        internal string DataTypeName(string documentName) =>
            $"{self.GeneratedName(documentName)}Data";

        /// <summary>
        /// 明示された名前設定を優先して、生成コード上の名前を決定します。
        /// </summary>
        /// <param name="sourceName">文書内で使用されている元の名前。</param>
        /// <returns>生成コード上で使用する名前。</returns>
        internal string GeneratedName(string sourceName) =>
            self.NameMappings.GetValueOrDefault(sourceName)
                ?? sourceName.ToCSharpIdentifier();
    }
}
