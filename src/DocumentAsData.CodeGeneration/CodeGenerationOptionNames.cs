namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// コード生成設定から生成後の名前を決定する処理を提供します。
/// </summary>
static class CodeGenerationOptionNames
{
    extension(CodeGenerationOptions self)
    {
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
