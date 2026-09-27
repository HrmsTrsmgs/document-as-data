using System.Globalization;

namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Documentラッパー生成で使用するテンプレート部品です。
/// </summary>
static partial class DocumentWrapperComponents
{
    /// <summary>
    /// 生成するC#ソースファイル全体を表すテンプレート部品です。
    /// </summary>
    /// <param name="filePath">生成元のWord文書のパス。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <param name="document">コード生成元の文書。</param>
    /// <returns>名前空間と文書型・データ型の宣言を含むC#コード。</returns>
    internal static string SourceFile(
        string filePath,
        CodeGenerationOptions options,
        Document document) =>
        $$"""
        using System;
        using System.IO;
        using System.Linq;
        using Marimo.DocumentAsData;

        namespace {{options.Namespace}};

        {{DocumentDeclaration(filePath, options, document)}}

        {{DataDeclaration(filePath, options, document)}}
        """;

    /// <summary>
    /// 型付き文書の宣言を生成します。ファイルの構成とは分け、Open・項目検証・データ入出力をまとめます。
    /// </summary>
    /// <param name="filePath">生成元のWord文書のパス。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <param name="document">コード生成元の文書。</param>
    /// <returns>型付き文書のC#コード。</returns>
    internal static string DocumentDeclaration(
        string filePath,
        CodeGenerationOptions options,
        Document document)
    {
        var documentName = Path.GetFileNameWithoutExtension(filePath);
        var typeName = options.DocumentTypeName(documentName);
        var dataTypeName = options.DataTypeName(documentName);

        return
            $$"""
            /// <summary>
            /// Word文書「{{documentName.EscapeAmpersands()}}」を型付きで表します。
            /// </summary>
            public partial class {{typeName}} : Document
            {
                {{typeName}}(string filePath)
                    : base(filePath)
                {
                }

                /// <summary>
                /// 指定したファイルパスのWord文書を型付きで開きます。
                /// </summary>
                /// <param name="filePath">開くWord文書のファイルパス。</param>
                /// <returns>開いた型付きWord文書。</returns>
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// 文字列項目は同名のMERGEFIELDまたは文字列Content Controlで読み取れます。
                /// </exception>
                public static new {{typeName}} Open(string filePath)
                {
            {{OpenMethodBody(typeName, "filePath")}}
                }

                {{typeName}}(Stream stream)
                    : base(stream)
                {
                }

                /// <summary>
                /// 指定したStream上のWord文書を型付きで開きます。
                /// </summary>
                /// <param name="stream">Word文書を保持するStream。</param>
                /// <returns>開いた型付きWord文書。</returns>
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// 文字列項目は同名のMERGEFIELDまたは文字列Content Controlで読み取れます。
                /// </exception>
                public static new {{typeName}} Open(Stream stream)
                {
            {{OpenMethodBody(typeName, "stream")}}
                }

            {{Validation.Declaration(document)}}

                /// <summary>
                /// Word文書全体のデータを読み込みます。
                /// </summary>
                /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
                public {{dataTypeName}} Read() =>
                    base.Read<{{dataTypeName}}>();

                /// <summary>
                /// Word文書全体のデータを置換します。
                /// </summary>
                /// <param name="data">文書へ書き込むデータ。</param>
                public void Replace({{dataTypeName}} data) =>
                    base.Replace(data);
            {{ForEach([
                .. from mergeField in TopLevel(document.MergeFields)
                   select MergeFieldPropertyDeclaration(mergeField, options),
                .. from contentControl in TopLevel(document.ContentControls)
                   select TextContentControlPropertyDeclaration(
                       contentControl,
                       options),
                .. from checkBox in TopLevel(document.CheckBoxes)
                   select CheckBoxPropertyDeclaration(checkBox, options),
                .. from datePicker in TopLevel(document.DatePickers)
                   select DatePickerPropertyDeclaration(datePicker, options),
                .. from section in document.RepeatingSections
                   select RepeatingSectionDeclaration(section, options)
            ])}}
            }
            """;
    }

    /// <summary>
    /// 繰り返しセクションの明細データ型と、型付きで列挙・置換するプロパティを生成します。
    /// 最初の明細を、生成するプロパティの見本として使用します。
    /// </summary>
    /// <param name="section">生成元の繰り返しセクション。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>明細データ型と列挙プロパティのC#コード。</returns>
    static string RepeatingSectionDeclaration(RepeatingSection section, CodeGenerationOptions options)
    {
        var dataTypeName = options.DataTypeName(section.Tag);
        var templateItem = section.Items[0];

        return $$"""

            /// <summary>
            /// 繰り返しセクション「{{section.Tag.EscapeAmpersands().EscapeLineBreaks()}}」を型付きで取得します。
            /// </summary>
            public RepeatingSection<{{dataTypeName}}> {{options.GeneratedName(section.Tag)}} =>
                new(RepeatingSections[{{StringLiteral(section.Tag)}}]);

            /// <summary>
            /// 繰り返しセクション「{{section.Tag.EscapeAmpersands().EscapeLineBreaks()}}」の明細1件分のデータです。
            /// </summary>
            public class {{dataTypeName}}
            {
            {{DataPropertyDeclarations(
                [
                    .. templateItem.ContentControls.Select(it => it.Tag),
                    .. templateItem.MergeFields.Select(it => it.Name)
                ],
                templateItem.CheckBoxes,
                templateItem.DatePickers,
                options)}}
            }
        """;
    }

    /// <summary>
    /// パス版とStream版に共通するOpenの本体を生成します。
    /// 検証に失敗した文書は解放し、成功した文書だけを呼び出し元へ返します。
    /// </summary>
    /// <param name="typeName">生成するDocument型の名前。</param>
    /// <param name="argumentName">文書を開くためのパスまたはStreamの引数名。</param>
    /// <returns>生成メソッド内のインデントを含むC#コード。</returns>
    static string OpenMethodBody(string typeName, string argumentName) =>
        $$"""
                var document = new {{typeName}}({{argumentName}});

                try
                {
                    document.ValidateRequiredItems();
                    return document;
                }
                catch
                {
                    document.Dispose();
                    throw;
                }
        """;

    /// <summary>
    /// Word文書全体のデータを表す型の宣言を生成します。
    /// </summary>
    /// <param name="filePath">生成元のWord文書のパス。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <param name="document">コード生成元の文書。</param>
    /// <returns>文書データ型のC#コード。</returns>
    internal static string DataDeclaration(
        string filePath,
        CodeGenerationOptions options,
        Document document)
    {
        var documentName = Path.GetFileNameWithoutExtension(filePath);

        return
            $$"""
            /// <summary>
            /// Word文書「{{documentName.EscapeAmpersands()}}」のデータを表します。
            /// </summary>
            public partial class {{options.DataTypeName(documentName)}}
            {
            {{DataPropertyDeclarations(
                [
                    .. TopLevel(document.MergeFields).Select(it => it.Name),
                    .. TopLevel(document.ContentControls).Select(it => it.Tag)
                ],
                TopLevel(document.CheckBoxes),
                TopLevel(document.DatePickers),
                options)}}
            }
            """;
    }

    /// <summary>
    /// 文書全体と明細のデータ型に共通する、値プロパティの並びを生成します。
    /// 文字列項目の順序と対象範囲は呼び出し元で決め、ここでは種類ごとの部品へ接続します。
    /// </summary>
    /// <param name="textItemNames">生成順に並べた文字列項目の元名。</param>
    /// <param name="checkBoxes">生成対象のチェックボックス。</param>
    /// <param name="datePickers">生成対象の日付選択。</param>
    /// <param name="options">名前変換の設定。</param>
    /// <returns>データ型の値プロパティ宣言。</returns>
    static string DataPropertyDeclarations(
        IEnumerable<string> textItemNames,
        IEnumerable<CheckBox> checkBoxes,
        IEnumerable<DatePicker> datePickers,
        CodeGenerationOptions options) =>
        ForEach([
            .. from name in textItemNames
               select DataTextPropertyDeclaration(name, options),
            .. from checkBox in checkBoxes
               select DataCheckBoxPropertyDeclaration(checkBox, options),
            .. from datePicker in datePickers
               select DataDatePickerPropertyDeclaration(datePicker, options)
        ]);

    /// <summary>
    /// 文書直下の生成プロパティに含めるMERGEFIELDを選びます。
    /// 全体検索を使う検証・診断には、この選別を適用しません。
    /// </summary>
    /// <param name="items">文書全体のMERGEFIELD。</param>
    /// <returns>明細に属さないMERGEFIELD。</returns>
    static IEnumerable<MergeField> TopLevel(IEnumerable<MergeField> items) =>
        from item in items
        where !item.IsInRepeatingSection
        select item;

    /// <summary>
    /// 文書直下の生成プロパティに含めるContent Controlを選びます。
    /// 種類ごとに同じ所属条件を繰り返し記述しないための選別です。
    /// </summary>
    /// <typeparam name="T">生成対象のContent Controlの種類。</typeparam>
    /// <param name="items">文書全体のContent Control。</param>
    /// <returns>明細に属さないContent Control。</returns>
    static IEnumerable<T> TopLevel<T>(IEnumerable<T> items) where T : ContentControl =>
        from item in items
        where !item.IsInRepeatingSection
        select item;

    /// <summary>
    /// 文書データ型に、文字列項目を表すプロパティ宣言を生成します。
    /// </summary>
    /// <param name="itemName">プロパティとして公開する文書項目の名前。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>文字列項目プロパティのC#コード。</returns>
    internal static string DataTextPropertyDeclaration(
        string itemName,
        CodeGenerationOptions options)
    {
        var propertyName = options.GeneratedName(itemName);

        return $$"""

            /// <summary>
            /// 文書項目「{{itemName.EscapeAmpersands().EscapeLineBreaks()}}」の文字列を取得または設定します。
            /// </summary>
            {{DataItemNameAttributeDeclaration(itemName, propertyName)}}public string {{propertyName}} { get; set; } = "";
        """;
    }

    /// <summary>
    /// 文書データ型に、CheckBoxを表すプロパティ宣言を生成します。
    /// </summary>
    /// <param name="checkBox">プロパティとして公開するCheckBox。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>CheckBoxプロパティのC#コード。</returns>
    internal static string DataCheckBoxPropertyDeclaration(
        CheckBox checkBox,
        CodeGenerationOptions options)
    {
        var propertyName = options.GeneratedName(checkBox.Tag);

        return $$"""

            /// <summary>
            /// CheckBox「{{checkBox.Tag.EscapeLineFeed()}}」のチェック状態を取得または設定します。
            /// </summary>
            {{DataItemNameAttributeDeclaration(checkBox.Tag, propertyName)}}public bool {{propertyName}} { get; set; }
        """;
    }

    /// <summary>
    /// 文書データ型に、DatePickerを表すプロパティ宣言を生成します。
    /// </summary>
    /// <param name="datePicker">プロパティとして公開するDatePicker。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>DatePickerプロパティのC#コード。</returns>
    internal static string DataDatePickerPropertyDeclaration(
        DatePicker datePicker,
        CodeGenerationOptions options)
    {
        var propertyName = options.GeneratedName(datePicker.Tag);

        return $$"""

            /// <summary>
            /// DatePicker「{{datePicker.Tag.EscapeLineFeed()}}」の日時を取得または設定します。
            /// </summary>
            {{DataItemNameAttributeDeclaration(datePicker.Tag, propertyName)}}public DateTimeOffset {{propertyName}} { get; set; }
        """;
    }

    /// <summary>
    /// 生成Dataプロパティ名が文書項目名と異なる場合に、ReadとReplaceへ元名を伝える属性を生成します。
    /// 書式文字はコンパイル後の識別子では無視されるため、ソース上で同名でも属性を付けます。
    /// </summary>
    /// <param name="itemName">文書内の項目名またはTag。</param>
    /// <param name="propertyName">生成するC#プロパティ名。</param>
    /// <returns>属性と次行のインデント。元名を補う必要がない場合は空文字列。</returns>
    static string DataItemNameAttributeDeclaration(string itemName, string propertyName) =>
        itemName == propertyName
            && !propertyName.Any(it => char.GetUnicodeCategory(it) == UnicodeCategory.Format)
            ? ""
            : $"[DocumentItemName({StringLiteral(itemName)})]{Environment.NewLine}    ";

    /// <summary>
    /// 生成Document型から指定したMERGEFIELDを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="mergeField">プロパティとして公開するMERGEFIELD。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>MERGEFIELDプロパティのC#コード。</returns>
    internal static string MergeFieldPropertyDeclaration(
        MergeField mergeField,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// MERGEFIELD「{{mergeField.Name.EscapeLineFeed()}}」の文字列を取得または設定します。
            /// </summary>
            /// <remarks>
            /// 読み取りと書き込みはMERGEFIELDを優先し、存在しない場合に同名の文字列Content Controlを使用します。
            /// </remarks>
            public string {{options.GeneratedName(mergeField.Name)}}
            {
                get => (MergeFields.SingleOrDefault(it => !it.IsInRepeatingSection && it.Name == {{StringLiteral(mergeField.Name)}}) as IDocumentTextItem
                    ?? ContentControls[{{StringLiteral(mergeField.Name)}}]).Text;
                set => (MergeFields.SingleOrDefault(it => !it.IsInRepeatingSection && it.Name == {{StringLiteral(mergeField.Name)}}) as IDocumentTextItem
                    ?? ContentControls[{{StringLiteral(mergeField.Name)}}]).Text = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定したDatePickerを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="datePicker">プロパティとして公開するDatePicker。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>DatePickerプロパティのC#コード。</returns>
    internal static string DatePickerPropertyDeclaration(
        DatePicker datePicker,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// DatePicker「{{datePicker.Tag.EscapeLineFeed()}}」の日時を取得または設定します。
            /// </summary>
            public DateTimeOffset {{options.GeneratedName(datePicker.Tag)}}
            {
                get => DatePickers[{{StringLiteral(datePicker.Tag)}}].SelectedDateTime;
                set => DatePickers[{{StringLiteral(datePicker.Tag)}}].SelectedDateTime = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定したCheckBoxを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="checkBox">プロパティとして公開するCheckBox。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>CheckBoxプロパティのC#コード。</returns>
    internal static string CheckBoxPropertyDeclaration(
        CheckBox checkBox,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// CheckBox「{{checkBox.Tag.EscapeLineFeed()}}」のチェック状態を取得または設定します。
            /// </summary>
            public bool {{options.GeneratedName(checkBox.Tag)}}
            {
                get => CheckBoxes.Single(it => !it.IsInRepeatingSection && it.Tag == {{StringLiteral(checkBox.Tag)}}).IsChecked;
                set => CheckBoxes.Single(it => !it.IsInRepeatingSection && it.Tag == {{StringLiteral(checkBox.Tag)}}).IsChecked = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定した文字列Content Controlを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="contentControl">プロパティとして公開する文字列Content Control。</param>
    /// <param name="options">コード生成時に適用する設定。</param>
    /// <returns>文字列Content ControlプロパティのC#コード。</returns>
    internal static string TextContentControlPropertyDeclaration(
        TextContentControl contentControl,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// 文字列Content Control「{{contentControl.Tag.EscapeAmpersands().EscapeLineBreaks()}}」の文字列を取得または設定します。
            /// </summary>
            /// <remarks>
            /// 読み取りと書き込みは文字列Content Controlを優先し、存在しない場合に同名のMERGEFIELDを使用します。
            /// </remarks>
            public string {{options.GeneratedName(contentControl.Tag)}}
            {
                get => (ContentControls.SingleOrDefault(it => !it.IsInRepeatingSection && it.Tag == {{StringLiteral(contentControl.Tag)}}) as IDocumentTextItem
                    ?? MergeFields[{{StringLiteral(contentControl.Tag)}}]).Text;
                set => (ContentControls.SingleOrDefault(it => !it.IsInRepeatingSection && it.Tag == {{StringLiteral(contentControl.Tag)}}) as IDocumentTextItem
                    ?? MergeFields[{{StringLiteral(contentControl.Tag)}}]).Text = value;
            }
        """;

    /// <summary>
    /// 複数のテンプレート部品を、生成ソース上の行単位で連結します。
    /// </summary>
    internal static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    /// <summary>
    /// 生成コード内へ埋め込む文字列リテラルを作ります。
    /// </summary>
    internal static string StringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").EscapeLineBreaks() + "\"";

    extension(string self)
    {
        /// <summary>
        /// LFを可視のエスケープ表記へ変換し、生成ソースの行が分断されるのを防ぎます。
        /// </summary>
        string EscapeLineFeed() =>
            self.Replace("\n", "\\n");

        /// <summary>
        /// CRとLFを可視のエスケープ表記へ変換します。
        /// 文字列リテラルでは、元のバックスラッシュをエスケープした後に呼び出します。
        /// </summary>
        string EscapeLineBreaks() =>
            self.Replace("\r", "\\r").EscapeLineFeed();

        /// <summary>
        /// 名前に含まれるアンパサンドがXMLコメントでエンティティとして解釈されるのを防ぎます。
        /// </summary>
        string EscapeAmpersands() =>
            self.Replace("&", "&amp;");
    }
}
