using System.Reflection;
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
    /// 派生した型付き文書から、指定したDOCXファイルを開きます。
    /// </summary>
    /// <param name="filePath">開くDOCXファイルのパス。</param>
    protected Document(string filePath)
        : this(OpenWorkingCopy(filePath))
    {
    }

    /// <summary>
    /// 派生した型付き文書から、指定したストリーム上のDOCX文書を開きます。
    /// </summary>
    /// <param name="stream">DOCX文書を格納したストリーム。</param>
    /// <remarks>
    /// 文書への変更はストリームへ書き戻しますが、ストリーム自体は閉じません。
    /// </remarks>
    protected Document(Stream stream)
        : this(Packaging.WordprocessingDocument.Open(stream, true))
    {
    }

    /// <summary>
    /// ファイルパス版で開いた文書と、その保存元および作業コピーを所有します。
    /// </summary>
    /// <param name="workingCopy">開いた文書と、その生存期間中に保持するストリーム。</param>
    Document((
        Packaging.WordprocessingDocument Document,
        Stream SourceStream,
        Stream WorkingStream) workingCopy)
        : this(
            workingCopy.Document,
            workingCopy.SourceStream,
            workingCopy.WorkingStream)
    {
    }

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
        DatePickers = new(this);
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
    /// <exception cref="DocumentMappingException">
    /// DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    public T Read<T>()
    {
        var data = Activator.CreateInstance<T>();

        foreach (var property in typeof(T).GetProperties())
        {
            var name = property.GetCustomAttribute<DocumentItemAttribute>()?.Name
                ?? property.Name;

            property.SetValue(data, ReadValue(name, property.PropertyType));
        }

        return data;
    }

    /// <summary>
    /// 同じ名前のContent ControlまたはMERGEFIELDから値を読み込みます。
    /// </summary>
    /// <param name="name">読み込む名前。</param>
    /// <param name="propertyType">読み込み先のプロパティ型。</param>
    /// <returns>文書から読み込んだ値。</returns>
    object ReadValue(string name, Type propertyType) =>
        propertyType == typeof(DateTimeOffset)
            ? (FindDatePicker(name) ?? throw new DocumentMappingException()).SelectedDateTime
            : (FindValueTarget(name) ?? throw new InvalidOperationException()).Text;

    /// <summary>
    /// 指定したオブジェクトのプロパティを、同じ名前のContent ControlまたはMERGEFIELDへ書き込みます。
    /// </summary>
    /// <typeparam name="T">文書へ書き込むデータの型。</typeparam>
    /// <param name="data">文書へ書き込むデータ。</param>
    /// <exception cref="DocumentMappingException">
    /// プロパティに対応する文書項目が存在しないか、
    /// 複数のプロパティが同じ文書項目に対応するか、
    /// DocumentItem属性を指定したプロパティにpublicなgetterがないか、
    /// プロパティの型に対応していない場合。
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// 同じ名前のContent ControlとMERGEFIELDが両方に存在する場合。
    /// </exception>
    public void Replace<T>(T data)
    {
        var mappings = (
            from property in typeof(T).GetProperties()
            let attribute = property.GetCustomAttribute<DocumentItemAttribute>()
            where attribute is not null || property.GetMethod?.IsPublic == true
            select new
            {
                Property = property,
                ItemName = attribute?.Name ?? property.Name,
                IsExplicitlyMapped = attribute is not null
            }
        ).ToArray();

        if (mappings.Any(it =>
            it.IsExplicitlyMapped &&
            it.Property.GetMethod?.IsPublic != true))
        {
            throw new DocumentMappingException();
        }

        if (mappings.Any(it => !IsSupportedPropertyType(it.Property.PropertyType)))
        {
            throw new DocumentMappingException();
        }

        if (mappings.Select(it => it.ItemName).Distinct().Count() != mappings.Length)
        {
            throw new DocumentMappingException();
        }

        foreach (var mapping in mappings)
        {
            ReplaceValue(
                mapping.ItemName,
                mapping.Property.GetValue(data)!);
        }
    }

    /// <summary>
    /// オブジェクトとの対応付けで扱えるプロパティ型かを取得します。
    /// </summary>
    /// <param name="type">確認するプロパティ型。</param>
    /// <returns>対応している型の場合は<c>true</c>。</returns>
    static bool IsSupportedPropertyType(Type type) =>
        type == typeof(string) ||
        type == typeof(DateTimeOffset);

    /// <summary>
    /// 同じ名前の文書項目へ値を書き込みます。
    /// </summary>
    /// <param name="name">書き込む名前。</param>
    /// <param name="value">書き込む値。</param>
    /// <exception cref="DocumentMappingException">対応する文書項目が存在しない場合。</exception>
    void ReplaceValue(string name, object value)
    {
        if (value is DateTimeOffset dateTime)
        {
            var datePicker = FindDatePicker(name) ?? throw new DocumentMappingException();

            datePicker.SelectedDateTime = dateTime;
            return;
        }

        var target = FindValueTarget(name) ?? throw new DocumentMappingException();

        target.Text = (string)value;
    }

    /// <summary>
    /// 同じ名前のContent ControlとMERGEFIELDから、一件だけある読み書き対象を取得します。
    /// </summary>
    /// <param name="name">取得する名前。</param>
    /// <returns>取得した文字列データ項目。存在しない場合はnull。</returns>
    /// <exception cref="InvalidOperationException">同じ名前の対象が複数存在する場合。</exception>
    DocumentTextItem? FindValueTarget(string name)
    {
        IEnumerable<DocumentTextItem> targets =
        [
            .. from contentControl in ContentControls
               where contentControl.Tag == name
               select contentControl,
            .. from mergeField in MergeFields
               where mergeField.Name == name
               select mergeField
        ];

        return targets.SingleOrDefault();
    }

    /// <summary>
    /// 同じTagの日付選択Content Controlから、一件だけある対象を取得します。
    /// </summary>
    /// <param name="tag">取得するTag。</param>
    /// <returns>取得した日付選択Content Control。存在しない場合はnull。</returns>
    /// <exception cref="InvalidOperationException">同じTagの対象が複数存在する場合。</exception>
    DatePicker? FindDatePicker(string tag) =>
        (
            from datePicker in DatePickers
            where datePicker.Tag == tag
            select datePicker
        ).SingleOrDefault();

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
