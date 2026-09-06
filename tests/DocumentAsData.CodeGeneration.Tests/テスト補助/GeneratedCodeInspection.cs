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
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    internal static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        DocumentWrapperGenerator.GenerateSources(filePath, configure);

    /// <summary>
    /// 製品コードのコード生成診断APIを呼び出します。
    /// </summary>
    /// <param name="filePath">診断対象のWord文書。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>コード生成前に検出された診断情報。</returns>
    internal static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        DocumentWrapperGenerator.GenerateDiagnostics(filePath, configure);

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

    extension(Type self)
    {
        /// <summary>
        /// 生成型のpublicなstaticメソッドを呼び出し、指定した型として結果を取得します。
        /// </summary>
        /// <typeparam name="T">呼び出し結果を扱う型。</typeparam>
        /// <param name="methodName">呼び出すメソッド名。</param>
        /// <param name="arguments">メソッドへ渡す引数。</param>
        /// <returns>指定した型として扱う呼び出し結果。</returns>
        internal T InvokeStaticMethod<T>(
            string methodName,
            params object?[] arguments) =>
            self.InvokeMember(
                methodName,
                BindingFlags.Public |
                    BindingFlags.Static |
                    BindingFlags.InvokeMethod,
                binder: null,
                target: null,
                arguments) is T result
                    ? result
                    : throw new InvalidOperationException(methodName);
    }
}
