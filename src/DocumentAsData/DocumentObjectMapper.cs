using System.Reflection;

using Marimo.DocumentAsData.CodeGeneration;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の名前付き項目とオブジェクトのプロパティを対応付けます。
/// </summary>
sealed class DocumentObjectMapper
{
    /// <summary>
    /// 対応付けの対象範囲にある文字列Content Controlです。
    /// </summary>
    readonly ContentControlCollection contentControls;

    /// <summary>
    /// 対応付けの対象範囲にあるMERGEFIELDです。
    /// </summary>
    readonly MergeFieldCollection mergeFields;

    /// <summary>
    /// 対応付けの対象範囲にあるチェックボックスです。
    /// </summary>
    readonly CheckBoxCollection checkBoxes;

    /// <summary>
    /// 対応付けの対象範囲にある日付選択Content Controlです。
    /// </summary>
    readonly DatePickerCollection datePickers;

    /// <summary>
    /// 指定した文書に対するオブジェクト対応付けを作成します。
    /// </summary>
    /// <param name="document">名前付き項目の読み書き対象となる文書。</param>
    internal DocumentObjectMapper(Document document)
        : this(document.ContentControls, document.MergeFields, document.CheckBoxes, document.DatePickers)
    {
    }

    /// <summary>
    /// 同じ検索範囲に属する項目を、オブジェクト対応付けの対象にします。
    /// </summary>
    /// <param name="contentControls">文字列Content Control。</param>
    /// <param name="mergeFields">MERGEFIELD。</param>
    /// <param name="checkBoxes">チェックボックス。</param>
    /// <param name="datePickers">日付選択Content Control。</param>
    internal DocumentObjectMapper(
        ContentControlCollection contentControls,
        MergeFieldCollection mergeFields,
        CheckBoxCollection checkBoxes,
        DatePickerCollection datePickers)
    {
        this.contentControls = contentControls;
        this.mergeFields = mergeFields;
        this.checkBoxes = checkBoxes;
        this.datePickers = datePickers;
    }

    /// <summary>
    /// 文書内のContent ControlまたはMERGEFIELDを、指定した型のpublicなインスタンスプロパティへ対応付けて読み込みます。
    /// 対応属性がない場合、文字列Content Control、CheckBox、日付選択のTagとMERGEFIELD名をC#識別子へ変換して照合します。
    /// </summary>
    /// <typeparam name="T">文書のデータを読み込む型。</typeparam>
    /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
    /// <exception cref="DocumentMappingException">
    /// プロパティの型に対応していないか、DateTimeOffsetプロパティに対応する日付選択Content Controlが存在しない場合。
    /// </exception>
    internal T Read<T>()
    {
        // 値型でも各SetValueが同じインスタンスを更新するよう、一度だけボックス化します。
        object? data = Activator.CreateInstance<T>();

        foreach (var (property, value) in
            from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            let name = property.GetCustomAttribute<DocumentItemNameAttribute>()?.Name
                ?? ResolveItemName(property.Name)
            select (property, ReadValue(name, property.PropertyType)))
        {
            property.SetValue(data, value);
        }

        // Tとして生成した値を戻します。nullの場合も、元のTが許すnullのまま返します。
        return (T)data!;
    }

    /// <summary>
    /// 同じ名前のContent ControlまたはMERGEFIELDから値を読み込みます。
    /// </summary>
    /// <param name="name">読み込む名前。</param>
    /// <param name="propertyType">読み込み先のプロパティ型。</param>
    /// <returns>文書から読み込んだ値。</returns>
    object ReadValue(string name, Type propertyType) =>
        propertyType == typeof(DateTimeOffset)
            ? (datePickers.FindByTag(name) ?? throw new DocumentMappingException())
                .SelectedDateTime
        : propertyType == typeof(bool)
            ? (checkBoxes.FindByTag(name) ?? throw new DocumentMappingException())
                .IsChecked
        : propertyType == typeof(string)
            ? (FindValueTarget(name) ?? throw new InvalidOperationException())
                .Text
        : throw new DocumentMappingException();

    /// <summary>
    /// 指定したオブジェクトのプロパティを、対応するContent ControlまたはMERGEFIELDへ書き込みます。
    /// 対応属性がない場合、文字列Content Control、CheckBox、日付選択のTagとMERGEFIELD名をC#識別子へ変換して照合します。
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
            from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
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
    /// 文字列Content Control、CheckBox、日付選択のTagとMERGEFIELD名をまとめてC#識別子へ変換して照合します。
    /// 該当しなければプロパティ名を使います。
    /// </summary>
    /// <param name="propertyName">読み書きするプロパティ名。</param>
    /// <returns>対応するTagまたはMERGEFIELD名、該当しなければ元のプロパティ名。</returns>
    string ResolveItemName(string propertyName)
    {
        IEnumerable<string> names =
        [
            .. contentControls.Select(it => it.Tag),
            .. mergeFields.Select(it => it.Name),
            .. checkBoxes.Select(it => it.Tag),
            .. datePickers.Select(it => it.Tag)
        ];

        return (
            from name in names
            where name.ToCSharpIdentifier() == propertyName
            select name
        ).SingleOrDefault() ?? propertyName;
    }

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
                (datePickers.FindByTag(name) ?? throw new DocumentMappingException())
                    .SelectedDateTime = dateTime;
                return;

            case bool isChecked:
                (checkBoxes.FindByTag(name) ?? throw new DocumentMappingException())
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
            .. from contentControl in contentControls
               where contentControl.Tag == name
               select contentControl,
            .. from mergeField in mergeFields
               where mergeField.Name == name
               select mergeField
        ];

        return targets.SingleOrDefault();
    }
}
