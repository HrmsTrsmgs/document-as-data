using Packaging = DocumentFormat.OpenXml.Packaging;

namespace Marimo.DocumentAsData;

/// <summary>
/// DOCXファイルとして開いた文書を表します。
/// </summary>
public class Document : IDisposable
{
    readonly Packaging.WordprocessingDocument document;

    Document(Packaging.WordprocessingDocument document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定したDOCXファイルを文書として開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(string filePath) =>
        new(Packaging.WordprocessingDocument.Open(filePath, true));

    /// <summary>
    /// 文書が使用しているファイルを閉じます。
    /// </summary>
    public void Close() =>
        document.Dispose();

    /// <summary>
    /// 文書が使用しているリソースを解放します。
    /// </summary>
    public void Dispose() =>
        Close();
}
