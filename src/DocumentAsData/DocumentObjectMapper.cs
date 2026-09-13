using System.Reflection;

using Marimo.DocumentAsData.CodeGeneration;

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
    /// 対応属性がない場合、文字列Content ControlとCheckBoxのTag、およびMERGEFIELD名はC#識別子へ変換して照合します。
    /// </summary>
    /// <typeparam name="T">文書のデータを読み込む型。</typeparam>
    /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
    /// <exception cref="DocumentMappingException">
    /// DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    internal T Read<T>()
    {
        var data = Activator.CreateInstance<T>();

        foreach (var (property, value) in
            from property in typeof(T).GetProperties()
            let name = property.GetCustomAttribute<DocumentItemNameAttribute>()?.Name
                ?? ResolveItemName(property.Name)
            select (property, ReadValue(name, property.PropertyType)))
        {
            property.SetValue(data, value);
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
            ? (FindDatePicker(name) ?? throw new DocumentMappingException())
                .SelectedDateTime
        : propertyType == typeof(bool)
            ? (FindCheckBox(name) ?? throw new DocumentMappingException())
                .IsChecked
        : (FindValueTarget(name) ?? throw new InvalidOperationException())
            .Text;

    /// <summary>
    /// 指定したオブジェクトのプロパティを、同じ名前のContent ControlまたはMERGEFIELDへ書き込みます。
    /// 対応属性がない場合、文字列Content ControlとCheckBoxのTag、およびMERGEFIELD名はC#識別子へ変換して照合します。
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
    internal void Replace<T>(T data)
    {
        var mappings = (
            from property in typeof(T).GetProperties()
            let attribute = property.GetCustomAttribute<DocumentItemNameAttribute>()
            where attribute is not null || property.GetMethod?.IsPublic == true
            select new
            {
                Property = property,
                ItemName = attribute?.Name ?? ResolveItemName(property.Name),
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
    /// 属性で名前を指定していないプロパティに対応する、元の文書項目名を取得します。
    /// 文字列Content ControlのTag、MERGEFIELD名、CheckBoxのTagの順にC#識別子へ変換して照合します。
    /// 該当しなければプロパティ名を使います。
    /// </summary>
    /// <param name="propertyName">読み書きするプロパティ名。</param>
    /// <returns>対応するTagまたはMERGEFIELD名、該当しなければ元のプロパティ名。</returns>
    string ResolveItemName(string propertyName) =>
        (
            from contentControl in document.ContentControls
            where contentControl.Tag.ToCSharpIdentifier() == propertyName
            select contentControl.Tag
        ).SingleOrDefault() ?? (
            from mergeField in document.MergeFields
            where mergeField.Name.ToCSharpIdentifier() == propertyName
            select mergeField.Name
        ).SingleOrDefault() ?? (
            from checkBox in document.CheckBoxes
            where checkBox.Tag.ToCSharpIdentifier() == propertyName
            select checkBox.Tag
        ).SingleOrDefault() ?? propertyName;

    /// <summary>
    /// オブジェクトとの対応付けで扱えるプロパティ型かを取得します。
    /// </summary>
    /// <param name="type">確認するプロパティ型。</param>
    /// <returns>対応している型の場合は<c>true</c>。</returns>
    static bool IsSupportedPropertyType(Type type) =>
        type == typeof(string) ||
        type == typeof(DateTimeOffset) ||
        type == typeof(bool);

    /// <summary>
    /// 同じ名前の文書項目へ値を書き込みます。
    /// </summary>
    /// <param name="name">書き込む名前。</param>
    /// <param name="value">書き込む値。</param>
    /// <exception cref="DocumentMappingException">対応する文書項目が存在しない場合。</exception>
    void ReplaceValue(string name, object value)
    {
        switch (value)
        {
            case DateTimeOffset dateTime:
                (FindDatePicker(name) ?? throw new DocumentMappingException())
                    .SelectedDateTime = dateTime;
                return;

            case bool isChecked:
                (FindCheckBox(name) ?? throw new DocumentMappingException())
                    .IsChecked = isChecked;
                return;

            default:
                (FindValueTarget(name) ?? throw new DocumentMappingException())
                    .Text = (string)value;
                return;
        }
    }

    /// <summary>
    /// 同じ名前のContent ControlとMERGEFIELDから、一件だけある読み書き対象を取得します。
    /// </summary>
    /// <param name="name">取得する名前。</param>
    /// <returns>取得した文字列データ項目。存在しない場合はnull。</returns>
    /// <exception cref="InvalidOperationException">同じ名前の対象が複数存在する場合。</exception>
    IDocumentTextItem? FindValueTarget(string name)
    {
        IEnumerable<IDocumentTextItem> targets =
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

    /// <summary>
    /// 同じTagのCheckBoxから、一件だけある対象を取得します。
    /// </summary>
    /// <param name="tag">取得するTag。</param>
    /// <returns>取得したCheckBox。存在しない場合はnull。</returns>
    /// <exception cref="InvalidOperationException">同じTagの対象が複数存在する場合。</exception>
    CheckBox? FindCheckBox(string tag) =>
        (
            from checkBox in document.CheckBoxes
            where checkBox.Tag == tag
            select checkBox
        ).SingleOrDefault();
}
