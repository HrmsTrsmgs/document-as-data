using System.Reflection;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の名前付き項目とオブジェクトのプロパティを対応付けます。
/// </summary>
sealed class DocumentObjectMapper
{
    /// <summary>
    /// 名前付き項目の読み書き対象となる文書です。
    /// </summary>
    readonly Document document;

    /// <summary>
    /// 指定した文書に対するオブジェクト対応付けを作成します。
    /// </summary>
    /// <param name="document">名前付き項目の読み書き対象となる文書。</param>
    internal DocumentObjectMapper(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 文書内のContent ControlまたはMERGEFIELDを、指定した型のプロパティへ対応付けて読み込みます。
    /// </summary>
    /// <typeparam name="T">文書のデータを読み込む型。</typeparam>
    /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
    /// <exception cref="DocumentMappingException">
    /// DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    internal T Read<T>()
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
    internal void Replace<T>(T data)
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
            .. from contentControl in document.ContentControls
               where contentControl.Tag == name
               select contentControl,
            .. from mergeField in document.MergeFields
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
            from datePicker in document.DatePickers
            where datePicker.Tag == tag
            select datePicker
        ).SingleOrDefault();
}
