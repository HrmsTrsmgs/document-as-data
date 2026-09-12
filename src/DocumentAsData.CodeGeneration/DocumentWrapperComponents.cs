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
            using System.Linq;
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
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// 文字列項目は同名のMERGEFIELDまたは文字列Content Controlで読み取れます。
                /// </exception>
                public static new {{typeName}} Open(string filePath)
                {
                    var document = new {{typeName}}(filePath);

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
                }

                {{typeName}}(System.IO.Stream stream)
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
                public static new {{typeName}} Open(System.IO.Stream stream)
                {
                    var document = new {{typeName}}(stream);

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
                }

                /// <summary>
                /// 生成元の文書にあったMERGEFIELD、文字列Content Control、CheckBox、DatePickerが存在することを検証します。
                /// 文字列項目は同名のMERGEFIELDと文字列Content Controlを読み取り先として認めます。
                /// 文書の生成、返却、解放は呼び出し元のOpenで行います。
                /// </summary>
                /// <exception cref="DocumentMappingException">
                /// 生成元の文書に対応する項目が不足している場合。
                /// </exception>
                void ValidateRequiredItems()
                {
            {{RequiredItemsValidation(
                [.. from mergeField in document.MergeFields select mergeField.Name],
                "[.. MergeFields.Select(it => it.Name), .. ContentControls.Select(it => it.Tag)]",
                "MERGEFIELD")}}
            {{RequiredItemsValidation(
                [.. from contentControl in document.ContentControls select contentControl.Tag],
                "[.. ContentControls.Select(it => it.Tag), .. MergeFields.Select(it => it.Name)]",
                "文字列Content Control")}}
            {{RequiredItemsValidation(
                [.. from checkBox in document.CheckBoxes select checkBox.Tag],
                "CheckBoxes.Select(it => it.Tag)",
                "CheckBox")}}
            {{RequiredItemsValidation(
                [.. from datePicker in document.DatePickers select datePicker.Tag],
                "DatePickers.Select(it => it.Tag)",
                "DatePicker")}}
                }

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
    /// 必要な名前の不足を検出する検査コードを生成します。対象が0件なら生成しません。
    /// </summary>
    /// <param name="requiredNames">生成元の文書に存在する、その種類の項目名またはTag。</param>
    /// <param name="actualNamesExpression">開いた文書から同じ種類の名前を列挙するC#式。</param>
    /// <returns>不足時に例外を投げるC#コード。対象が0件なら空文字列。</returns>
    static string RequiredItemsValidation(string[] requiredNames, string actualNamesExpression) =>
        requiredNames.Length == 0
            ? ""
            : $$"""
                      if (new[] { {{string.Join(", ", requiredNames.Select(StringLiteral))}} }
                          .Except({{actualNamesExpression}})
                          .Any())
                      {
                          throw new DocumentMappingException();
                      }
              """;

    /// <summary>
    /// 不足した名前と種類を例外メッセージへ含める検査コードを生成します。
    /// </summary>
    /// <param name="requiredNames">生成元の文書に存在する、その種類の項目名またはTag。</param>
    /// <param name="actualNamesExpression">開いた文書から同じ種類の名前を列挙するC#式。</param>
    /// <param name="itemKind">メッセージに表示する項目の種類。</param>
    /// <returns>詳細メッセージ付きの検査コード。対象が0件なら空文字列。</returns>
    static string RequiredItemsValidation(
        string[] requiredNames,
        string actualNamesExpression,
        string itemKind) =>
        requiredNames.Length == 0
            ? ""
            : $$"""
                      {
                          var missingNames = new[] { {{string.Join(", ", requiredNames.Select(StringLiteral))}} }
                              .Except({{actualNamesExpression}})
                              .ToArray();

                          if (missingNames.Length > 0)
                          {
                              throw new DocumentMappingException(
                                  {{StringLiteral($"{itemKind}が不足しています: ")}} + string.Join(", ", missingNames));
                          }
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

        return $$"""

            /// <summary>
            /// 文書項目「{{itemName}}」の文字列を取得または設定します。
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
            /// CheckBox「{{checkBox.Tag}}」のチェック状態を取得または設定します。
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
            /// DatePicker「{{datePicker.Tag}}」の日時を取得または設定します。
            /// </summary>
            {{DataItemNameAttributeDeclaration(datePicker.Tag, propertyName)}}public System.DateTimeOffset {{propertyName}} { get; set; }
        """;
    }

    /// <summary>
    /// 生成Dataプロパティ名が文書項目名と異なる場合に、ReadとReplaceへ元名を伝える属性を生成します。
    /// </summary>
    /// <param name="itemName">文書内の項目名またはTag。</param>
    /// <param name="propertyName">生成するC#プロパティ名。</param>
    /// <returns>属性と次行のインデント。同名の場合は空文字列。</returns>
    static string DataItemNameAttributeDeclaration(string itemName, string propertyName) =>
        itemName == propertyName
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
            /// MERGEFIELD「{{mergeField.Name}}」の文字列を取得または設定します。
            /// </summary>
            /// <remarks>
            /// 読み取りと書き込みでは同名の文字列Content Controlも対象にします。
            /// </remarks>
            public string {{options.GeneratedName(mergeField.Name)}}
            {
                get => Enumerable.Single<IDocumentTextItem>(
                    [
                        .. from item in ContentControls
                           where item.Tag == {{StringLiteral(mergeField.Name)}}
                           select item,
                        .. from item in MergeFields
                           where item.Name == {{StringLiteral(mergeField.Name)}}
                           select item
                    ]).Text;
                set => Enumerable.Single<IDocumentTextItem>(
                    [
                        .. from item in ContentControls
                           where item.Tag == {{StringLiteral(mergeField.Name)}}
                           select item,
                        .. from item in MergeFields
                           where item.Name == {{StringLiteral(mergeField.Name)}}
                           select item
                    ]).Text = value;
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
            /// <remarks>
            /// 読み取りと書き込みでは同名のMERGEFIELDも対象にします。
            /// </remarks>
            public string {{options.GeneratedName(contentControl.Tag)}}
            {
                get => Enumerable.Single<IDocumentTextItem>(
                    [
                        .. from item in ContentControls
                           where item.Tag == {{StringLiteral(contentControl.Tag)}}
                           select item,
                        .. from item in MergeFields
                           where item.Name == {{StringLiteral(contentControl.Tag)}}
                           select item
                    ]).Text;
                set => Enumerable.Single<IDocumentTextItem>(
                    [
                        .. from item in ContentControls
                           where item.Tag == {{StringLiteral(contentControl.Tag)}}
                           select item,
                        .. from item in MergeFields
                           where item.Name == {{StringLiteral(contentControl.Tag)}}
                           select item
                    ]).Text = value;
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
