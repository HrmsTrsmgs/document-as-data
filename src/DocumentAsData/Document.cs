using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Validation = DocumentFormat.OpenXml.Validation;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// DOCXファイルとして開いた文書を表します。
/// </summary>
public class Document : IDisposable
{
    /// <summary>
    /// 読み取り・書き込み対象のOpen XML文書です。
    /// </summary>
    readonly Packaging.WordprocessingDocument document;

    /// <summary>
    /// ファイルパス版で開いた保存元を、文書の生存期間中に束縛するストリームです。
    /// Stream版ではnullです。
    /// </summary>
    readonly Stream? sourceStream;

    /// <summary>
    /// 保存元を変更せずに編集するための作業コピーです。
    /// Stream版ではnullです。
    /// </summary>
    readonly Stream? workingStream;

    /// <summary>
    /// ファイルパス版で使う保存元と作業コピーを、Open XML文書と同じ生存期間で解放できるよう所有します。
    /// Stream版ではストリームを所有しません。
    /// </summary>
    /// <param name="document">読み取り・書き込み対象のOpen XML文書。</param>
    /// <param name="sourceStream">保存元を束縛するストリーム。</param>
    /// <param name="workingStream">編集対象の作業コピー。</param>
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
        CheckBoxes = new(this);
    }

    /// <summary>
    /// 本文に含まれるOOXML要素を文書順に取得します。
    /// </summary>
    internal IEnumerable<OpenXmlElement> Elements =>
        document.MainDocumentPart?.Document?.Descendants() ?? [];

    /// <summary>
    /// 指定したDOCXファイルを、保存元を変更しない作業コピーとして開きます。
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

    /// <summary>
    /// 指定した保存元ストリームを複製し、編集可能な文書として開きます。
    /// </summary>
    /// <param name="sourceStream">複製するDOCX文書のストリーム。</param>
    /// <returns>保存元と作業コピーを所有する文書。</returns>
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
    /// 指定したDOCXファイルを、必要に応じて検証し、保存元を変更しない作業コピーとして開きます。
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
    /// <remarks>
    /// 文書への変更はストリームへ書き戻しますが、ストリーム自体は閉じません。
    /// </remarks>
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
    /// <remarks>
    /// 文書への変更はストリームへ書き戻しますが、ストリーム自体は閉じません。
    /// </remarks>
    public static Document Open(Stream stream, bool validate) =>
        ValidateIfRequested(Open(stream), validate);

    /// <summary>
    /// 指定された場合だけ文書をOpen XMLとして検証します。
    /// </summary>
    /// <param name="opened">検証対象の文書。</param>
    /// <param name="validate">文書を検証する場合は<c>true</c>。</param>
    /// <returns>検証が不要または検証に成功した文書。</returns>
    /// <exception cref="InvalidDataException">Open XML検証エラーがある場合。</exception>
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
    /// 文書内のContent ControlまたはMERGEFIELDを、指定した型のプロパティへ対応付けて読み込みます。
    /// </summary>
    /// <typeparam name="T">文書のデータを読み込む型。</typeparam>
    /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
    public T Read<T>()
    {
        var data = Activator.CreateInstance<T>();

        foreach (var property in typeof(T).GetProperties())
        {
            property.SetValue(data, ReadValue(property.Name));
        }

        return data;
    }

    /// <summary>
    /// 同じ名前のContent ControlまたはMERGEFIELDから値を読み込みます。
    /// </summary>
    /// <param name="name">読み込む名前。</param>
    /// <returns>文書から読み込んだ値。</returns>
    string ReadValue(string name) =>
        ContentControls.Any(it => it.Tag == name)
            ? ContentControls[name].Value
            : MergeFields[name].Value;

    /// <summary>
    /// 指定したオブジェクトのプロパティを、同じ名前のContent ControlまたはMERGEFIELDへ書き込みます。
    /// </summary>
    /// <typeparam name="T">文書へ書き込むデータの型。</typeparam>
    /// <param name="data">文書へ書き込むデータ。</param>
    public void Replace<T>(T data)
    {
        foreach (var property in typeof(T).GetProperties())
        {
            ReplaceValue(property.Name, (string)property.GetValue(data)!);
        }
    }

    /// <summary>
    /// 同じ名前のContent ControlまたはMERGEFIELDへ値を書き込みます。
    /// </summary>
    /// <param name="name">書き込む名前。</param>
    /// <param name="value">書き込む値。</param>
    void ReplaceValue(string name, string value)
    {
        if (ContentControls.Any(it => it.Tag == name))
        {
            ContentControls[name].Value = value;
            return;
        }

        MergeFields[name].Value = value;
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
    /// 文書内のチェックボックスを取得するコレクションを取得します。
    /// </summary>
    public CheckBoxCollection CheckBoxes { get; }

    /// <summary>
    /// 保存元を変更せず、文書を別のDOCXファイルとして保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    public void SaveAs(string filePath)
    {
        using var savedDocument = document.Clone(filePath);
    }

    /// <summary>
    /// 保存元を変更せず、文書が使用しているファイルを閉じます。
    /// </summary>
    public void Close()
    {
        document.Dispose();
        workingStream?.Dispose();
        sourceStream?.Dispose();
    }

    /// <summary>
    /// 保存元を変更せず、文書が使用しているリソースを解放します。
    /// </summary>
    public void Dispose() =>
        Close();
}
