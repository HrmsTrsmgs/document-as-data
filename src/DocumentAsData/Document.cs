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
    /// SDK文書と入出力リソースを同じ生存期間で管理します。
    /// </summary>
    readonly DocumentSession session;

    /// <summary>
    /// 文書内の名前付き項目とオブジェクトのプロパティを対応付けます。
    /// </summary>
    readonly DocumentObjectMapper objectMapper;

    /// <summary>
    /// 派生した型付き文書から、指定したDOCXファイルを開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    protected Document(string filePath)
        : this(DocumentSession.Open(filePath))
    {
    }

    /// <summary>
    /// 派生した型付き文書から、指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">現在位置からDOCX文書を読み取れるストリーム。</param>
    /// <exception cref="InvalidDataException">現在位置以降にデータがない場合。</exception>
    /// <remarks>
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
    /// 読み取り専用・非シークの入力も使用できます。現在位置を文書の先頭として扱います。
    /// </remarks>
    protected Document(Stream stream)
        : this(DocumentSession.Open(stream))
    {
    }

    /// <summary>
    /// 開いた文書のセッションを保持し、名前付き項目とオブジェクトの対応付けを初期化します。
    /// </summary>
    /// <param name="session">SDK文書と入出力リソースを所有するセッション。</param>
    Document(DocumentSession session)
    {
        this.session = session;
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
        session.Document.MainDocumentPart?.Document?.Descendants() ?? [];

    /// <summary>
    /// 指定したDOCXファイルを、保存元を変更しない作業コピーとして開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    /// <returns>開いた文書。</returns>
    public static Document Open(string filePath) =>
        new(filePath);

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
    /// <param name="stream">現在位置からDOCX文書を読み取れるストリーム。</param>
    /// <returns>開いた文書。</returns>
    /// <exception cref="InvalidDataException">現在位置以降にデータがない場合。</exception>
    /// <remarks>
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
    /// 読み取り専用・非シークの入力も使用できます。現在位置を文書の先頭として扱います。
    /// </remarks>
    public static Document Open(Stream stream) =>
        new(stream);

    /// <summary>
    /// 指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">現在位置からDOCX文書を読み取れるストリーム。</param>
    /// <param name="validate">開く文書をOpen XMLとして検証する場合は<c>true</c>。</param>
    /// <returns>開いた文書。</returns>
    /// <exception cref="InvalidDataException">
    /// 現在位置以降にデータがない、または<paramref name="validate" />が<c>true</c>で文書にOpen XML検証エラーがある場合。
    /// </exception>
    /// <remarks>
    /// 元ストリームの内容は変更せず、ストリーム自体も閉じません。Saveは使用できません。
    /// 読み取り専用・非シークの入力も使用できます。現在位置を文書の先頭として扱います。
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
        if (!validate || !new Validation.OpenXmlValidator().Validate(opened.session.Document).Any())
        {
            return opened;
        }

        opened.Dispose();
        throw new InvalidDataException();
    }

    /// <summary>
    /// 文書内のContent ControlまたはMERGEFIELDを、指定した型のプロパティへ対応付けて読み込みます。
    /// </summary>
    /// <remarks>
    /// DocumentItemName属性に指定した元名を優先します。属性がない場合は文書内の名前・Tagを
    /// C#識別子へ変換してプロパティ名と照合し、該当がなければプロパティ名をそのまま使います。
    /// 読み込み先はpublicなインスタンスプロパティに限定し、staticプロパティは変更しません。
    /// </remarks>
    /// <typeparam name="T">文書のデータを読み込む型。</typeparam>
    /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
    /// <exception cref="DocumentMappingException">
    /// プロパティの型に対応していないか、DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    public T Read<T>() =>
        objectMapper.Read<T>();

    /// <summary>
    /// 指定したオブジェクトのpublicなインスタンスプロパティを、対応するContent ControlまたはMERGEFIELDへ書き込みます。
    /// staticプロパティは書き込み元として使用しません。
    /// </summary>
    /// <remarks>
    /// DocumentItemName属性に指定した元名を優先します。属性がない場合は文書内の名前・Tagを
    /// C#識別子へ変換してプロパティ名と照合し、該当がなければプロパティ名をそのまま使います。
    /// </remarks>
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
    public void Save() =>
        session.Save(GetType());

    /// <summary>
    /// 保存元を変更せず、文書を別のDOCXファイルとして保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    public void SaveAs(string filePath) =>
        session.SaveAs(filePath);

    /// <summary>
    /// 未保存の変更を保存せずに文書を閉じます。呼び出し側のStreamは閉じません。
    /// </summary>
    public void Close() =>
        session.Dispose();

    /// <summary>
    /// 未保存の変更を保存せずに文書を閉じて所有リソースを解放します。
    /// </summary>
    public void Dispose() =>
        Close();

    /// <summary>
    /// SDK文書の保存と生存期間を管理します。Streamの入出力契約とは別に、
    /// 保存元ファイルの束縛と作業領域の所有権を引き受けます。
    /// </summary>
    sealed class DocumentSession : IDisposable
    {
        /// <summary>
        /// 保存元を変更せずにSDKが読み書きする作業領域です。
        /// </summary>
        readonly Stream workingStream;

        /// <summary>
        /// このセッションが所有する保存元のStreamです。作業領域とは別の実体です。
        /// パス版では元ファイルの束縛を維持し、Save時だけ開き直します。
        /// Stream版では保存元を所有しないためnullです。
        /// </summary>
        Stream? sourceStream;

        /// <summary>
        /// Saveで上書きする元ファイルの絶対パスです。Stream版はnullで、上書き保存を提供しません。
        /// </summary>
        readonly string? filePath;

        /// <summary>
        /// 二重解放と、解放後の保存によるファイルの再束縛を防ぎます。
        /// </summary>
        bool closed;

        /// <summary>
        /// 作業領域上でSDK文書を開き、領域の所有権を引き受けます。
        /// 開けなかった場合も作業領域を解放します。
        /// </summary>
        /// <param name="workingStream">SDKが読み書きする、このセッション所有のStream。</param>
        /// <param name="sourceStream">このセッションが所有する保存元のStream。借用Stream版ではnull。</param>
        /// <param name="filePath">保存元の絶対パス。借用Stream版ではnull。</param>
        DocumentSession(Stream workingStream, Stream? sourceStream, string? filePath)
        {
            this.workingStream = workingStream;
            this.sourceStream = sourceStream;
            this.filePath = filePath;
            try
            {
                Document = Packaging.WordprocessingDocument.Open(workingStream, true, new OpenSettings { AutoSave = false });
            }
            catch
            {
                workingStream.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 項目の読み書きとOOXML検証に使うSDK文書です。解放と保存はセッションが管理します。
        /// </summary>
        internal Packaging.WordprocessingDocument Document { get; }

        /// <summary>
        /// 保存元ファイルを束縛し、先に作業コピーを作って開きます。
        /// Saveで元ファイルを閉じても、SDKの読み取り先が失われないようにします。
        /// </summary>
        /// <param name="filePath">開くDOCXファイルのパス。</param>
        /// <returns>元ファイルと作業領域を所有するセッション。</returns>
        internal static DocumentSession Open(string filePath)
        {
            filePath = Path.GetFullPath(filePath);
            var sourceStream = File.OpenRead(filePath);
            try
            {
                return new(CreateWorkingCopy(sourceStream), sourceStream, filePath);
            }
            catch
            {
                sourceStream.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 借りたStreamへの書き込みをコピーへ隔離して開きます。元Streamは所有しません。
        /// </summary>
        /// <param name="source">呼び出し側が所有する、現在位置からDOCXを読み取れるStream。</param>
        /// <returns>SDK文書とコピー切り替え用のラッパーを所有するセッション。</returns>
        /// <exception cref="InvalidDataException">入力の現在位置以降に文書データがない場合。</exception>
        internal static DocumentSession Open(Stream source)
        {
            var workingStream = new CopyOnWriteStream(source);
            // SDKは空Streamも開けるため、文書がない入力はここで拒否します。
            // 非シーク入力のために作成したコピーだけを解放し、借りた元Streamは閉じません。
            if (workingStream.Length == 0)
            {
                workingStream.Dispose();
                throw new InvalidDataException();
            }

            return new(workingStream, null, null);
        }

        /// <summary>
        /// 保存元ファイルから独立した編集領域を作り、先頭から読み取れる状態にします。
        /// </summary>
        /// <param name="source">先頭に位置する保存元ファイルのStream。</param>
        /// <returns>元ファイルの内容を複製した、拡張可能な編集領域。</returns>
        static MemoryStream CreateWorkingCopy(Stream source)
        {
            var copy = new MemoryStream();
            try
            {
                source.CopyTo(copy);
                copy.Position = 0;
                return copy;
            }
            catch
            {
                copy.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 元ファイルの束縛を保存中だけ解除し、保存処理が終わったら再び束縛します。
        /// </summary>
        /// <param name="documentType">解放済み例外で報告する、公開文書オブジェクトの実行時型。</param>
        /// <exception cref="NotSupportedException">呼び出し側から借りたStreamの場合。</exception>
        /// <exception cref="ObjectDisposedException">所有するファイルが既に閉じられている場合。</exception>
        internal void Save(Type documentType)
        {
            if (filePath is null)
            {
                throw new NotSupportedException();
            }

            ObjectDisposedException.ThrowIf(closed, documentType);
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
        /// SDKの複製を閉じることで、戻る時点で保存先のDOCXを完成させ、束縛を残しません。
        /// </summary>
        /// <param name="filePath">保存先ファイルのパス。</param>
        internal void SaveAs(string filePath)
        {
            using var savedDocument = Document.Clone(filePath);
        }

        /// <summary>
        /// SDK文書、作業領域、元ファイルの順で解放します。未保存の変更は書き戻しません。
        /// </summary>
        public void Dispose()
        {
            if (closed)
            {
                return;
            }

            try
            {
                Document.Dispose();
            }
            finally
            {
                workingStream.Dispose();
                sourceStream?.Dispose();
                closed = true;
            }
        }

        /// <summary>
        /// SDKの書き込みを拡張可能なコピーへ隔離し、元Streamの内容を保護します。
        /// シーク可能な入力は最初の書き込み時、非シーク入力は読み取り前にコピーします。
        /// 元Streamは借りるだけで、このラッパーは閉じません。文書の保存やファイルの束縛は扱いません。
        /// </summary>
        sealed class CopyOnWriteStream : Stream
        {
            /// <summary>
            /// コピーへの切り替え前の読み取り元です。所有権は呼び出し側に残します。
            /// </summary>
            readonly Stream source;

            /// <summary>
            /// Open時の位置を文書の先頭とし、前置データを読み書き対象から除きます。
            /// </summary>
            readonly long sourceOffset;

            /// <summary>
            /// SDKによる変更を保持する作業領域です。非シーク入力では読み取り前に作成します。
            /// </summary>
            MemoryStream? workingCopy;

            /// <summary>
            /// 入力の現在位置を記録します。シーク不可の入力は、SDKがランダムアクセスできるよう先にコピーします。
            /// </summary>
            /// <param name="source">呼び出し側が所有する、現在位置から文書を読み取れるStream。</param>
            internal CopyOnWriteStream(Stream source)
            {
                this.source = source;
                if (source.CanSeek)
                {
                    sourceOffset = source.Position;
                }
                else
                {
                    CreateWorkingCopy();
                }
            }

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
            public override long Length => workingCopy?.Length ?? source.Length - sourceOffset;

            /// <inheritdoc />
            public override long Position
            {
                get => workingCopy?.Position ?? source.Position - sourceOffset;
                set => Current.Position = workingCopy is null ? sourceOffset + value : value;
            }

            /// <inheritdoc />
            public override int Read(byte[] buffer, int offset, int count) =>
                Current.Read(buffer, offset, count);

            /// <inheritdoc />
            public override long Seek(long offset, SeekOrigin origin) =>
                workingCopy is not null
                    ? workingCopy.Seek(offset, origin)
                    : source.Seek(origin == SeekOrigin.Begin ? sourceOffset + offset : offset, origin) - sourceOffset;

            /// <inheritdoc />
            public override void Write(byte[] buffer, int offset, int count) =>
                CreateWorkingCopy().Write(buffer, offset, count);

            /// <inheritdoc />
            public override void SetLength(long value) =>
                CreateWorkingCopy().SetLength(value);

            /// <inheritdoc />
            public override void Flush() => workingCopy?.Flush();

            /// <summary>
            /// 文書先頭からの読み書き位置を保ってコピーします。非シーク入力の位置は参照しません。
            /// 以後は同じ作業領域を使い、元Streamへ書き戻しません。
            /// </summary>
            /// <returns>このラッパーが所有する拡張可能な作業領域。</returns>
            MemoryStream CreateWorkingCopy()
            {
                if (workingCopy is null)
                {
                    var position = source.CanSeek ? Position : 0;
                    if (source.CanSeek)
                    {
                        source.Position = sourceOffset;
                    }
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
}
