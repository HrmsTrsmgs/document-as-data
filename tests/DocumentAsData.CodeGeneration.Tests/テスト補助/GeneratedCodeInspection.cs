using System.Reflection;
using Marimo.DocumentAsData.CodeGeneration;

namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// コード生成テストで、生成結果を観測するための入口を提供します。
/// </summary>
static class GeneratedCodeInspection
{
    /// <summary>
    /// テスト内で生成コードをコンパイルするときの既定の名前空間です。
    /// </summary>
    internal const string NamespaceName = "Generated";

    /// <summary>
    /// 指定したWord文書から生成されるC#ソースコードを取得します。
    /// </summary>
    /// <param name="filePath">コード生成元のWord文書。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    internal static string[] GenerateSources(string filePath) =>
        DocumentWrapperGenerator.GenerateSources(filePath);

    /// <summary>
    /// 生成済みソースコードをコンパイルします。
    /// </summary>
    /// <param name="sources">コンパイルするC#ソースコード。</param>
    /// <returns>生成コードをコンパイルしたアセンブリ。</returns>
    internal static Assembly AssemblyFrom(IEnumerable<string> sources) =>
        GeneratedSourceCompiler.Compile(sources);

    extension(Assembly self)
    {
        /// <summary>
        /// コンパイルしたアセンブリから、指定した生成型を取得します。
        /// </summary>
        /// <param name="typeName">既定名前空間を除いた生成型名。</param>
        /// <returns>指定した生成型。</returns>
        internal Type GeneratedType(string typeName) =>
            self.GetType($"{NamespaceName}.{typeName}")
                ?? throw new InvalidOperationException(typeName);
    }
}
