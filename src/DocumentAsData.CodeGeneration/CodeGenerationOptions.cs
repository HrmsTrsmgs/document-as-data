namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Word文書からC#ラッパーコードを生成するときの設定を表します。
/// </summary>
public sealed class CodeGenerationOptions
{
    /// <summary>設定を省略した場合に生成型を配置する名前空間です。</summary>
    internal const string DefaultNamespace = "Generated";

    /// <summary>
    /// 生成するC#型を配置する名前空間を取得または設定します。
    /// </summary>
    public string Namespace { get; set; } = DefaultNamespace;

    /// <summary>
    /// 文書内の名前から生成後のC#名への対応表を取得または設定します。
    /// </summary>
    public Dictionary<string, string> NameMappings { get; set; } = [];
}
