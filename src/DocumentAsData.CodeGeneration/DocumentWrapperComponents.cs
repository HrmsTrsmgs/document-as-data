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
        var typeName = $"{documentName.ToCSharpIdentifier()}Document";

        return
            $$"""
            using Marimo.DocumentAsData;

            namespace {{options.Namespace}};

            public partial class {{typeName}} : Document
            {
                {{typeName}}(string filePath)
                    : base(filePath)
                {
                }

                public static new {{typeName}} Open(string filePath) =>
                    new(filePath);

                {{typeName}}(System.IO.Stream stream)
                    : base(stream)
                {
                }

                public static new {{typeName}} Open(System.IO.Stream stream) =>
                    new(stream);
            {{ForEach([
                .. from mergeField in document.MergeFields
                   select MergeFieldPropertyDeclaration(mergeField),
                .. from contentControl in document.ContentControls
                   select TextContentControlPropertyDeclaration(contentControl),
                .. from checkBox in document.CheckBoxes
                   select CheckBoxPropertyDeclaration(checkBox),
                .. from datePicker in document.DatePickers
                   select DatePickerPropertyDeclaration(datePicker)
            ])}}
            }
            """;
    }

    /// <summary>
    /// 生成Document型から指定したMERGEFIELDを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="mergeField">プロパティとして公開するMERGEFIELD。</param>
    /// <returns>MERGEFIELDプロパティのC#コード。</returns>
    internal static string MergeFieldPropertyDeclaration(MergeField mergeField) =>
        $$"""

            public string {{mergeField.Name.ToCSharpIdentifier()}}
            {
                get => MergeFields["{{mergeField.Name}}"].Text;
                set => MergeFields["{{mergeField.Name}}"].Text = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定したDatePickerを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="datePicker">プロパティとして公開するDatePicker。</param>
    /// <returns>DatePickerプロパティのC#コード。</returns>
    internal static string DatePickerPropertyDeclaration(DatePicker datePicker) =>
        $$"""

            public System.DateTimeOffset {{datePicker.Tag}}
            {
                get => DatePickers["{{datePicker.Tag}}"].SelectedDateTime;
                set => DatePickers["{{datePicker.Tag}}"].SelectedDateTime = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定したCheckBoxを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="checkBox">プロパティとして公開するCheckBox。</param>
    /// <returns>CheckBoxプロパティのC#コード。</returns>
    internal static string CheckBoxPropertyDeclaration(CheckBox checkBox) =>
        $$"""

            public bool {{checkBox.Tag}}
            {
                get => CheckBoxes["{{checkBox.Tag}}"].IsChecked;
                set => CheckBoxes["{{checkBox.Tag}}"].IsChecked = value;
            }
        """;

    /// <summary>
    /// 生成Document型から指定した文字列Content Controlを取得するプロパティ宣言を生成します。
    /// </summary>
    /// <param name="contentControl">プロパティとして公開する文字列Content Control。</param>
    /// <returns>文字列Content ControlプロパティのC#コード。</returns>
    internal static string TextContentControlPropertyDeclaration(
        TextContentControl contentControl) =>
        $$"""

            public string {{contentControl.Tag.ToCSharpIdentifier()}}
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
