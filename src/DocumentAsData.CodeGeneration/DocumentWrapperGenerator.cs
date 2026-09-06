namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Word文書から、DocumentAsDataの型付きラッパーコードを生成します。
/// </summary>
public static class DocumentWrapperGenerator
{
    /// <summary>
    /// 指定したWord文書からC#ソースコードを生成します。
    /// </summary>
    /// <param name="filePath">生成元のWord文書のパス。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    public static string[] GenerateSources(string filePath)
    {
        var typeName = $"{Path.GetFileNameWithoutExtension(filePath)}Document";

        return
        [
            $$"""
            using Marimo.DocumentAsData;

            namespace Generated;

            public class {{typeName}} : Document
            {
                {{typeName}}(string filePath)
                    : base(filePath)
                {
                }
            }
            """
        ];
    }
}
