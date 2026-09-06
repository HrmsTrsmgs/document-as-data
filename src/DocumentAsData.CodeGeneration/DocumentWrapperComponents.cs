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
            """;
    }

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
                get => MergeFields["{{mergeField.Name}}"].Text;
                set => MergeFields["{{mergeField.Name}}"].Text = value;
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
                get => DatePickers["{{datePicker.Tag}}"].SelectedDateTime;
                set => DatePickers["{{datePicker.Tag}}"].SelectedDateTime = value;
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
                get => CheckBoxes["{{checkBox.Tag}}"].IsChecked;
                set => CheckBoxes["{{checkBox.Tag}}"].IsChecked = value;
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
                get => ContentControls["{{contentControl.Tag}}"].Text;
                set => ContentControls["{{contentControl.Tag}}"].Text = value;
            }
        """;

    /// <summary>
    /// 複数のテンプレート部品を、生成ソース上の行単位で連結します。
    /// </summary>
    internal static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);
}
