using Packaging = DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Validation = DocumentFormat.OpenXml.Validation;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// DOCXファイルとして開いた文書を表します。
/// </summary>
public class Document : IDisposable
{
    readonly Packaging.WordprocessingDocument document;
    readonly Stream? sourceStream;
    readonly Stream? workingStream;

    Document(
        Packaging.WordprocessingDocument document,
        Stream? sourceStream = null,
        Stream? workingStream = null)
    {
        this.document = document;
        this.sourceStream = sourceStream;
        this.workingStream = workingStream;
        MergeFields = new(this);
        ContentControls = new(this);
    }

    internal IEnumerable<OpenXmlElement> Elements =>
        document.MainDocumentPart?.Document?.Descendants() ?? [];

    /// <summary>
    /// 指定したDOCXファイルを文書として開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(string filePath)
    {
        var sourceStream = File.OpenRead(filePath);
        try
        {
            return OpenWorkingCopy(sourceStream);
        }
        catch
        {
            sourceStream.Dispose();
            throw;
        }
    }

    static Document OpenWorkingCopy(Stream sourceStream)
    {
        var workingStream = new MemoryStream();
        try
        {
            sourceStream.CopyTo(workingStream);
            workingStream.Position = 0;
            return new(
                Packaging.WordprocessingDocument.Open(workingStream, true),
                sourceStream,
                workingStream);
        }
        catch
        {
            workingStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 指定したDOCXファイルを文書として開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <param name="validate">開く文書をOpen XMLとして検証する場合は<c>true</c>。</param>
    /// <returns>開いた文書。</returns>
    /// <exception cref="InvalidDataException">
    /// <paramref name="validate" />が<c>true</c>で、文書にOpen XML検証エラーがある場合。
    /// </exception>
    public static Document Open(string filePath, bool validate) =>
        ValidateIfRequested(Open(filePath), validate);

    /// <summary>
    /// 指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">DOCX文書を格納したストリーム。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(Stream stream) =>
        new(Packaging.WordprocessingDocument.Open(stream, true));

    /// <summary>
    /// 指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">DOCX文書を格納したストリーム。</param>
    /// <param name="validate">開く文書をOpen XMLとして検証する場合は<c>true</c>。</param>
    /// <returns>開いた文書。</returns>
    /// <exception cref="InvalidDataException">
    /// <paramref name="validate" />が<c>true</c>で、文書にOpen XML検証エラーがある場合。
    /// </exception>
    public static Document Open(Stream stream, bool validate) =>
        ValidateIfRequested(Open(stream), validate);

    static Document ValidateIfRequested(Document opened, bool validate)
    {
        if (!validate || !new Validation.OpenXmlValidator().Validate(opened.document).Any())
        {
            return opened;
        }

        opened.Dispose();
        throw new InvalidDataException();
    }

    /// <summary>
    /// 文書内のMERGEFIELDを取得するコレクションを取得します。
    /// </summary>
    public MergeFieldCollection MergeFields { get; }

    /// <summary>
    /// 文書内のContent Controlを取得するコレクションを取得します。
    /// </summary>
    public ContentControlCollection ContentControls { get; }

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
    public void Close()
    {
        document.Dispose();
        workingStream?.Dispose();
        sourceStream?.Dispose();
    }

    /// <summary>
    /// 文書が使用しているリソースを解放します。
    /// </summary>
    public void Dispose() =>
        Close();
}
