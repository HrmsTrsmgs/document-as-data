using Packaging = DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml.Packaging;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

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
        MergeFields = new(this);
    }

    internal IEnumerable<Wordprocessing.SimpleField> SimpleFields =>
        document.MainDocumentPart?.Document?.Descendants<Wordprocessing.SimpleField>() ?? [];

    /// <summary>
    /// 指定したDOCXファイルを文書として開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(string filePath) =>
        new(Packaging.WordprocessingDocument.Open(filePath, true));

    /// <summary>
    /// 指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">DOCX文書を格納したストリーム。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(Stream stream) =>
        new(Packaging.WordprocessingDocument.Open(stream, true));

    /// <summary>
    /// 文書内のMERGEFIELDを取得するコレクションを取得します。
    /// </summary>
    public MergeFieldCollection MergeFields { get; }

    /// <summary>
    /// 文書内のContent Controlを取得するコレクションを取得します。
    /// </summary>
    public ContentControlCollection ContentControls =>
        throw new NotImplementedException();

    /// <summary>
    /// 文書を別のDOCXファイルとして保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    public void SaveAs(string filePath)
    {
        using var savedDocument = document.Clone(filePath);
    }

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
