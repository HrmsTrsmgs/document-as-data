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
    /// 文書内の名前付き項目とオブジェクトのプロパティを対応付けます。
    /// </summary>
    readonly DocumentObjectMapper objectMapper;

    /// <summary>
    /// ファイルパス版で開いた保存元を、文書の生存期間中に束縛するストリームです。
    /// Stream版ではnullです。
    /// </summary>
    Stream? sourceStream;

    /// <summary>
    /// Saveで上書きする元ファイルの絶対パスです。作業ディレクトリの変更に影響されません。
    /// Stream版ではnullです。
    /// </summary>
    readonly string? filePath;

    /// <summary>
    /// 保存元を変更せずに編集するための作業コピーです。
    /// Stream版では最初の書き込みでコピーを作るラッパーを保持します。
    /// </summary>
    readonly Stream? workingStream;

    /// <summary>
    /// 二重に解放をしないための終了状態です。
    /// </summary>
    bool closed;

    /// <summary>
    /// 派生した型付き文書から、指定したDOCXファイルを開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    protected Document(string filePath)
        : this(OpenWorkingCopy(filePath))
    {
        this.filePath = Path.GetFullPath(filePath);
    }

    /// <summary>
    /// 派生した型付き文書から、指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">DOCX文書を格納したストリーム。</param>
    /// <remarks>
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
    /// </remarks>
    protected Document(Stream stream)
        : this(OpenBorrowedStream(stream))
    {
    }

    /// <summary>
    /// 開いた文書と作業領域を所有します。保存元の所有権はファイル版だけが保持します。
    /// </summary>
    /// <param name="workingCopy">開いた文書と、その生存期間中に保持するストリーム。</param>
    Document((
        Packaging.WordprocessingDocument Document,
        Stream? SourceStream,
        Stream WorkingStream) workingCopy)
        : this(
            workingCopy.Document,
            workingCopy.SourceStream,
            workingCopy.WorkingStream)
    {
    }

    /// <summary>
    /// ファイルパス版で使う保存元と作業コピーを、Open XML文書と同じ生存期間で解放できるよう所有します。
    /// Stream版ではラッパーだけを所有し、呼び出し側のストリームは所有しません。
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
        DatePickers = new(this);
        objectMapper = new(this);
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
    public static Document Open(string filePath) =>
        new(filePath);

    /// <summary>
    /// 指定したDOCXファイルを開き、保存元を変更しない作業コピーを作成します。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <returns>開いた文書と、その生存期間中に保持するストリーム。</returns>
    static (
        Packaging.WordprocessingDocument Document,
        Stream SourceStream,
        Stream WorkingStream) OpenWorkingCopy(string filePath)
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
    /// <returns>開いた文書と、その生存期間中に保持するストリーム。</returns>
    static (
        Packaging.WordprocessingDocument Document,
        Stream SourceStream,
        Stream WorkingStream) OpenWorkingCopy(Stream sourceStream)
    {
        var workingStream = new MemoryStream();
        try
        {
            sourceStream.CopyTo(workingStream);
            workingStream.Position = 0;
            return (
                Packaging.WordprocessingDocument.Open(workingStream, true, new OpenSettings { AutoSave = false }),
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
    /// 借りたStreamへのSDKの書き込みを、作業コピーへ隔離して開きます。
    /// </summary>
    /// <param name="source">呼び出し側が所有するStream。</param>
    /// <returns>開いた文書と、この文書が所有するラッパー。</returns>
    static (
        Packaging.WordprocessingDocument Document,
        Stream? SourceStream,
        Stream WorkingStream) OpenBorrowedStream(Stream source)
    {
        var stream = new CopyOnWriteStream(source);
        try
        {
            return (
                Packaging.WordprocessingDocument.Open(stream, true, new OpenSettings { AutoSave = false }),
                null,
                stream);
        }
        catch
        {
            stream.Dispose();
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
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
    /// </remarks>
    public static Document Open(Stream stream) =>
        new(stream);

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
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
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
    /// <exception cref="DocumentMappingException">
    /// DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    public T Read<T>() =>
        objectMapper.Read<T>();

    /// <summary>
    /// 指定したオブジェクトのプロパティを、同じ名前のContent ControlまたはMERGEFIELDへ書き込みます。
    /// </summary>
    /// <typeparam name="T">文書へ書き込むデータの型。</typeparam>
    /// <param name="data">文書へ書き込むデータ。</param>
    /// <exception cref="DocumentMappingException">
    /// プロパティに対応する文書項目が存在しないか、
    /// 複数のプロパティが同じ文書項目に対応するか、
    /// DocumentItemName属性を指定したプロパティにpublicなgetterがないか、
    /// プロパティの型に対応していない場合。
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// 同じ名前のContent ControlとMERGEFIELDが両方に存在する場合。
    /// </exception>
    public void Replace<T>(T data) =>
        objectMapper.Replace(data);

    /// <summary>
    /// 文書内のMERGEFIELDを取得するコレクションを取得します。
    /// </summary>
    public MergeFieldCollection MergeFields { get; }

    /// <summary>
    /// 文書内の文字列Content Controlを取得するコレクションを取得します。
    /// </summary>
    public ContentControlCollection ContentControls { get; }

    /// <summary>
    /// 文書内のチェックボックスを取得するコレクションを取得します。
    /// </summary>
    public CheckBoxCollection CheckBoxes { get; }

    /// <summary>
    /// 文書内の日付選択Content Controlを取得するコレクションを取得します。
    /// </summary>
    public DatePickerCollection DatePickers { get; }

    /// <summary>
    /// 開いた元のファイルへ変更を保存します。Streamから開いた文書には使用できません。
    /// </summary>
    /// <exception cref="NotSupportedException">Streamから開いた文書の場合。</exception>
    /// <exception cref="ObjectDisposedException">ファイルから開いた文書が既に閉じられている場合。</exception>
    public void Save()
    {
        if (filePath is null)
        {
            throw new NotSupportedException();
        }

        ObjectDisposedException.ThrowIf(closed, this);

        sourceStream?.Dispose();
        try
        {
            SaveAs(filePath);
        }
        finally
        {
            sourceStream = File.OpenRead(filePath);
        }
    }

    /// <summary>
    /// 保存元を変更せず、文書を別のDOCXファイルとして保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    public void SaveAs(string filePath)
    {
        using var savedDocument = document.Clone(filePath);
    }

    /// <summary>
    /// 未保存の変更を保存せずに文書を閉じます。呼び出し側のStreamは閉じません。
    /// </summary>
    public void Close()
    {
        if (closed)
        {
            return;
        }

        try
        {
            document.Dispose();
        }
        finally
        {
            workingStream?.Dispose();
            sourceStream?.Dispose();
            closed = true;
        }
    }

    /// <summary>
    /// 未保存の変更を保存せずに文書を閉じて所有リソースを解放します。
    /// </summary>
    public void Dispose() =>
        Close();

    /// <summary>
    /// SDKの最初の書き込みで拡張可能なコピーへ切り替え、元Streamの内容を保護します。
    /// 元Streamは借りるだけで、このラッパーは閉じません。
    /// </summary>
    /// <param name="source">呼び出し側が所有する、読み取り・シーク可能な元Stream。</param>
    sealed class CopyOnWriteStream(Stream source) : Stream
    {
        /// <summary>
        /// SDKによる変更を保持する作業領域です。書き込み前は作成しません。
        /// </summary>
        MemoryStream? workingCopy;

        /// <summary>
        /// コピーへの切り替え前は元Stream、切り替え後は作業領域を読みます。
        /// </summary>
        Stream Current => workingCopy ?? source;

        /// <inheritdoc />
        public override bool CanRead => Current.CanRead;

        /// <inheritdoc />
        public override bool CanSeek => Current.CanSeek;

        /// <summary>
        /// 書き込み先には拡張可能なコピーを使います。
        /// </summary>
        public override bool CanWrite => true;

        /// <inheritdoc />
        public override long Length => Current.Length;

        /// <inheritdoc />
        public override long Position
        {
            get => Current.Position;
            set => Current.Position = value;
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) =>
            Current.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) =>
            Current.Seek(offset, origin);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) =>
            CreateWorkingCopy().Write(buffer, offset, count);

        /// <inheritdoc />
        public override void SetLength(long value) =>
            CreateWorkingCopy().SetLength(value);

        /// <inheritdoc />
        public override void Flush() => workingCopy?.Flush();

        /// <summary>
        /// 読み書き位置を保ったまま元Streamをコピーし、以後は同じ作業領域を使います。
        /// </summary>
        /// <returns>このラッパーが所有する拡張可能な作業領域。</returns>
        MemoryStream CreateWorkingCopy()
        {
            if (workingCopy is null)
            {
                var position = source.Position;
                source.Position = 0;
                var copy = new MemoryStream();
                try
                {
                    source.CopyTo(copy);
                    copy.Position = position;
                    workingCopy = copy;
                }
                catch
                {
                    copy.Dispose();
                    throw;
                }
            }

            return workingCopy;
        }

        /// <summary>
        /// 所有する作業コピーだけを解放し、呼び出し側のStreamは残します。
        /// </summary>
        /// <param name="disposing">マネージドリソースも解放する場合はtrue。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                workingCopy?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
