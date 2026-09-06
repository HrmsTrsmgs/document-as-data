using System.Reflection;
using System.Runtime.Loader;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// 生成ソースをメモリ上でコンパイルし、実行時型として検証するためのテスト補助です。
/// </summary>
static class GeneratedSourceCompiler
{
    /// <summary>
    /// 指定したC#ソースコードを、DocumentAsData本体を参照したアセンブリとしてコンパイルします。
    /// </summary>
    /// <param name="sources">コンパイルするC#ソースコード。</param>
    /// <returns>コンパイルしたアセンブリ。</returns>
    internal static Assembly Compile(IEnumerable<string> sources)
    {
        var syntaxTrees =
            from source in sources
            select CSharpSyntaxTree.ParseText(source);

        var compilation = CSharpCompilation.Create(
            $"DocumentAsData.Generated.{Guid.NewGuid():N}",
            syntaxTrees,
            References,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        using var stream = new MemoryStream();
        var result = compilation.Emit(stream);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                string.Join(
                    Environment.NewLine,
                    from diagnostic in result.Diagnostics
                    where diagnostic.Severity == DiagnosticSeverity.Error
                    select diagnostic.ToString()));
        }

        stream.Position = 0;
        return AssemblyLoadContext.Default.LoadFromStream(stream);
    }

    /// <summary>
    /// 生成ソースのコンパイルに必要な参照アセンブリです。
    /// </summary>
    static IEnumerable<MetadataReference> References =>
        (
            from path in
                TrustedPlatformAssemblyPaths
                    .Append(typeof(Document).Assembly.Location)
            where !string.IsNullOrEmpty(path)
            group path by path into paths
            select MetadataReference.CreateFromFile(paths.Key)
        );

    /// <summary>
    /// 現在の.NET実行環境が既定で参照できるアセンブリパスです。
    /// </summary>
    static IEnumerable<string> TrustedPlatformAssemblyPaths =>
        ((string?)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES"))
            ?.Split(Path.PathSeparator)
            ?? [];
}
