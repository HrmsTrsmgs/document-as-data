namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Documentラッパー生成で使用するテンプレート部品です。
/// </summary>
static class DocumentWrapperComponents
{
    /// <summary>
    /// 生成するC#ソースファイル全体を表すテンプレート部品です。
    /// </summary>
    internal static string SourceFile(
        string filePath,
        CodeGenerationOptions options,
        Document document)
    {
        var documentName = Path.GetFileNameWithoutExtension(filePath);
        var typeName = $"{options.GeneratedName(documentName)}Document";
        var dataTypeName = $"{options.GeneratedName(documentName)}Data";

        return
            $$"""
            using Marimo.DocumentAsData;

            namespace {{options.Namespace}};

            /// <summary>
            /// Word文書「{{documentName}}」を型付きで表します。
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
                public static new {{typeName}} Open(string filePath) =>
                    new(filePath);

                {{typeName}}(System.IO.Stream stream)
                    : base(stream)
                {
                }

                /// <summary>
                /// 指定したStream上のWord文書を型付きで開きます。
                /// </summary>
                /// <param name="stream">Word文書を保持するStream。</param>
                /// <returns>開いた型付きWord文書。</returns>
                public static new {{typeName}} Open(System.IO.Stream stream) =>
                    new(stream);

                /// <summary>
                /// Word文書全体のデータを読み込みます。
                /// </summary>
                /// <returns>文書内のデータを読み込んだオブジェクト。</returns>
                public new {{dataTypeName}} Read() =>
                    base.Read<{{dataTypeName}}>();

                /// <summary>
                /// Word文書全体のデータを置換します。
                /// </summary>
                /// <param name="data">文書へ書き込むデータ。</param>
                public new void Replace({{dataTypeName}} data) =>
                    base.Replace(data);
            {{ForEach([
                .. from mergeField in document.MergeFields
                   select MergeFieldPropertyDeclaration(mergeField, options),
                .. from contentControl in document.ContentControls
                   select TextContentControlPropertyDeclaration(
                       contentControl,
                       options),
                .. from checkBox in document.CheckBoxes
                   select CheckBoxPropertyDeclaration(checkBox, options),
                .. from datePicker in document.DatePickers
                   select DatePickerPropertyDeclaration(datePicker, options)
            ])}}
            }

            {{DataDeclaration(filePath, options, document)}}
            """;
    }

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
        var typeName = $"{options.GeneratedName(documentName)}Data";

        return
            $$"""
            /// <summary>
            /// Word文書「{{documentName}}」のデータを表します。
            /// </summary>
            public partial class {{typeName}}
            {
            {{ForEach([
                .. from mergeField in document.MergeFields
                   select DataTextPropertyDeclaration(mergeField.Name, options),
                .. from contentControl in document.ContentControls
                   select DataTextPropertyDeclaration(contentControl.Tag, options),
                .. from checkBox in document.CheckBoxes
                   select DataCheckBoxPropertyDeclaration(checkBox, options),
                .. from datePicker in document.DatePickers
                   select DataDatePickerPropertyDeclaration(datePicker, options)
            ])}}
            }
            """;
    }

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
        var attributeDeclaration = itemName == propertyName
            ? ""
            : $"[DocumentItem({StringLiteral(itemName)})]{Environment.NewLine}    ";

        return $$"""

            /// <summary>
            /// 文書項目「{{itemName}}」の文字列を取得または設定します。
            /// </summary>
            {{attributeDeclaration}}public string {{propertyName}} { get; set; } = "";
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
        var attributeDeclaration = checkBox.Tag == propertyName
            ? ""
            : $"[DocumentItem({StringLiteral(checkBox.Tag)})]{Environment.NewLine}    ";

        return $$"""

            /// <summary>
            /// CheckBox「{{checkBox.Tag}}」のチェック状態を取得または設定します。
            /// </summary>
            {{attributeDeclaration}}public bool {{propertyName}} { get; set; }
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
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// DatePicker「{{datePicker.Tag}}」の日時を取得または設定します。
            /// </summary>
            public System.DateTimeOffset {{options.GeneratedName(datePicker.Tag)}} { get; set; }
        """;

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
            /// MERGEFIELD「{{mergeField.Name}}」の文字列を取得または設定します。
            /// </summary>
            public string {{options.GeneratedName(mergeField.Name)}}
            {
                get => MergeFields[{{StringLiteral(mergeField.Name)}}].Text;
                set => MergeFields[{{StringLiteral(mergeField.Name)}}].Text = value;
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
            /// DatePicker「{{datePicker.Tag}}」の日時を取得または設定します。
            /// </summary>
            public System.DateTimeOffset {{options.GeneratedName(datePicker.Tag)}}
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
            /// CheckBox「{{checkBox.Tag}}」のチェック状態を取得または設定します。
            /// </summary>
            public bool {{options.GeneratedName(checkBox.Tag)}}
            {
                get => CheckBoxes[{{StringLiteral(checkBox.Tag)}}].IsChecked;
                set => CheckBoxes[{{StringLiteral(checkBox.Tag)}}].IsChecked = value;
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
            /// 文字列Content Control「{{contentControl.Tag}}」の文字列を取得または設定します。
            /// </summary>
            public string {{options.GeneratedName(contentControl.Tag)}}
            {
                get => ContentControls[{{StringLiteral(contentControl.Tag)}}].Text;
                set => ContentControls[{{StringLiteral(contentControl.Tag)}}].Text = value;
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
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
