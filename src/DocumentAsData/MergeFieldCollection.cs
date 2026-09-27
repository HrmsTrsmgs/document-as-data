using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを取得するコレクションを表します。
/// </summary>
public partial class MergeFieldCollection : IEnumerable<MergeField>
{
    /// <summary>
    /// 列挙方法をコレクションの生存期間中共有します。
    /// </summary>
    readonly MergeFieldReader reader;

    /// <summary>
    /// 文書の解析処理をコレクションの生存期間中共有できるようにします。
    /// </summary>
    /// <param name="document">MERGEFIELDを取得する文書。</param>
    internal MergeFieldCollection(Document document)
        : this(document, document.Elements)
    {
    }

    /// <summary>
    /// 指定した範囲の解析処理を保持します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="elements">MERGEFIELDを検索する範囲の要素列。</param>
    internal MergeFieldCollection(Document document, IEnumerable<OpenXmlElement> elements)
    {
        reader = new(document, elements);
    }

    /// <summary>
    /// 指定した名前のMERGEFIELDを取得します。
    /// </summary>
    /// <param name="name">取得するMERGEFIELDの名前。</param>
    /// <returns>指定した名前のMERGEFIELD。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前のMERGEFIELDが存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">指定した名前のMERGEFIELDが複数存在する場合。</exception>
    public MergeField this[string name] =>
        (
            from field in this
            where field.Name == name
            select field
        ).SingleOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// MERGEFIELDを列挙する列挙子を返します。
    /// </summary>
    /// <returns>MERGEFIELDを列挙する列挙子。</returns>
    public IEnumerator<MergeField> GetEnumerator() =>
        reader.Read().GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    /// <summary>
    /// MERGEFIELDの列挙にだけ必要なOOXML依存の解析処理を集約します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="elements">文書順に解析する対象範囲の要素列。</param>
    sealed partial class MergeFieldReader(Document document, IEnumerable<OpenXmlElement> elements)
    {
        /// <summary>
        /// 生成するMERGEFIELDの所属先です。
        /// </summary>
        readonly Document document = document;

        /// <summary>
        /// 文書全体または明細内など、MERGEFIELDを解析する範囲の要素列です。
        /// </summary>
        readonly IEnumerable<OpenXmlElement> elements = elements;

        /// <summary>
        /// 同じOOXML要素からは、検索範囲にかかわらず同じMERGEFIELDを返すために保持します。
        /// </summary>
        readonly OpenXmlElementCache<MergeField> cache = document.MergeFieldCache;

        /// <summary>
        /// 文書要素を先頭から読み取り、認識できたMERGEFIELDを文書順に返します。
        /// 複合フィールドの読み取り状態は列挙ごとに作り直します。
        /// </summary>
        /// <returns>文書内で認識できたMERGEFIELD。</returns>
        public IEnumerable<MergeField> Read()
        {
            var complexFieldReader = new ComplexFieldReader();

            return
                from element in elements
                from mergeField in GetMergeFields(element, complexFieldReader)
                select mergeField;
        }

        /// <summary>
        /// 一つのOOXML要素から得られる0件または1件のMERGEFIELDを、LINQで結合できる列挙として返します。
        /// </summary>
        /// <param name="element">読み取るOOXML要素。</param>
        /// <param name="complexFieldReader">複合フィールドの読み取り状態。</param>
        /// <returns>認識できた場合だけMERGEFIELDを含む列挙。</returns>
        IEnumerable<MergeField> GetMergeFields(
            OpenXmlElement element,
            ComplexFieldReader complexFieldReader)
        {
            if (TryGetMergeField(element, complexFieldReader, out var mergeField))
            {
                yield return mergeField;
            }
        }

        /// <summary>
        /// 単純フィールドはその場で変換し、それ以外の要素は複合フィールドの読み取り状態へ渡します。
        /// </summary>
        /// <param name="element">読み取るOOXML要素。</param>
        /// <param name="complexFieldReader">複合フィールドの読み取り状態。</param>
        /// <param name="mergeField">認識できたMERGEFIELD。認識できない場合はnull。</param>
        /// <returns>MERGEFIELDを認識できた場合はtrue。</returns>
        bool TryGetMergeField(
            OpenXmlElement element,
            ComplexFieldReader complexFieldReader,
            [NotNullWhen(true)] out MergeField? mergeField)
        {
            if (element is Wordprocessing.SimpleField simpleField)
            {
                return TryGetSimpleMergeField(simpleField, out mergeField);
            }

            if (complexFieldReader.TryRead(element, out var complexField))
            {
                return TryGetComplexMergeField(complexField, out mergeField);
            }

            mergeField = null;
            return false;
        }

        /// <summary>
        /// 単純フィールドの命令がMERGEFIELDなら、対応するMERGEFIELDを取得します。
        /// </summary>
        /// <param name="field">読み取る単純フィールド。</param>
        /// <param name="mergeField">認識できたMERGEFIELD。認識できない場合はnull。</param>
        /// <returns>MERGEFIELD命令だった場合はtrue。</returns>
        bool TryGetSimpleMergeField(
            Wordprocessing.SimpleField field,
            [NotNullWhen(true)] out MergeField? mergeField)
        {
            if (MergeFieldInstruction.TryParseName(
                field.Instruction?.Value,
                out var name))
            {
                mergeField =
                    cache.GetOrAdd(
                        field,
                        () => new(document, field, name));
                return true;
            }
            mergeField = null;
            return false;
        }

        /// <summary>
        /// 読み取り済みの複合フィールドの命令がMERGEFIELDなら、対応するMERGEFIELDを取得します。
        /// </summary>
        /// <param name="field">読み取りが完了した複合フィールド。</param>
        /// <param name="mergeField">認識できたMERGEFIELD。認識できない場合はnull。</param>
        /// <returns>MERGEFIELD命令だった場合はtrue。</returns>
        bool TryGetComplexMergeField(
            ComplexField field,
            [NotNullWhen(true)] out MergeField? mergeField)
        {
            if (MergeFieldInstruction.TryParseName(
                field.Instruction,
                out var name))
            {
                mergeField =
                    cache.GetOrAdd(
                        field.FieldStart,
                        () => new(
                            document,
                            name,
                            field.FieldStart,
                            field.ResultSeparator,
                            field.FieldEnd,
                            field.ResultElements));
                return true;
            }

            mergeField = null;
            return false;
        }

        /// <summary>
        /// OOXMLのフィールド命令を解釈し、MERGEFIELDかどうかとその名前を読み取ります。
        /// </summary>
        static partial class MergeFieldInstruction
        {
            /// <summary>
            /// フィールド命令がMERGEFIELDかを判定し、MERGEFIELD名を取り出します。
            /// </summary>
            /// <param name="fieldInstruction">OOXMLに保存されたフィールド命令。</param>
            /// <param name="name">MERGEFIELDだった場合の名前。それ以外の場合はnull。</param>
            /// <returns>MERGEFIELD命令だった場合はtrue。</returns>
            internal static bool TryParseName(
                string? fieldInstruction,
                [NotNullWhen(true)] out string? name)
            {
                const string fieldType = "MERGEFIELD";

                var instruction = fieldInstruction?.Trim();
                if (instruction is null ||
                    !instruction.StartsWith(
                        fieldType,
                        StringComparison.OrdinalIgnoreCase))
                {
                    name = null;
                    return false;
                }

                name = ParseName(instruction[fieldType.Length..]);
                return true;
            }

            /// <summary>
            /// MERGEFIELD命令の引数部分からMERGEFIELD名を取得します。
            /// </summary>
            /// <param name="fieldParameters">MERGEFIELDキーワードより後ろの命令。</param>
            /// <returns>引用符で囲まれた名前、または引用符なしの先頭トークン。</returns>
            static string ParseName(string fieldParameters)
            {
                var parameters = fieldParameters.Trim();
                if (!parameters.StartsWith('"'))
                {
                    var nameLength = parameters
                        .TakeWhile(it => !char.IsWhiteSpace(it))
                        .Count();
                    return parameters[..nameLength];
                }

                // \"は名前中の引用符として読み進め、エスケープされていない引用符で止めます。
                // フィールド命令で使うエスケープを外して、名前そのものを返します。
                return QuotedName().Match(parameters[1..])
                    .Value.Replace("\\\"", "\"");
            }

            /// <summary>
            /// 名前中のエスケープされた引用符を含め、閉じ引用符の手前までを取得します。
            /// </summary>
            /// <returns>引用符で囲まれた名前の内容を識別する正規表現。</returns>
            [GeneratedRegex("""^(?:\\"|[^"])*""")]
            private static partial Regex QuotedName();
        }

        /// <summary>
        /// OOXML要素を文書順に受け取り、開始から終了まで揃った複合フィールドを返します。
        /// 読み取り途中の命令と表示結果は、複合フィールドが終了するまで保持します。
        /// </summary>
        sealed class ComplexFieldReader
        {
            /// <summary>
            /// 現在読み取っている複合フィールドの開始要素です。読み取り中でない場合はnullです。
            /// </summary>
            Wordprocessing.FieldChar? fieldStart;

            /// <summary>
            /// 開始要素から結果開始要素までに分割して保存された命令文字列を保持します。
            /// </summary>
            string? fieldInstruction;

            /// <summary>
            /// 命令部分と表示結果を区切る要素です。表示結果の開始前はnullです。
            /// </summary>
            Wordprocessing.FieldChar? resultSeparator;

            /// <summary>
            /// 結果開始後に現れた表示用の文字列、タブ、改行要素を保持します。結果開始前はnullです。
            /// </summary>
            List<OpenXmlElement>? resultElements;

            /// <summary>
            /// OOXML要素を一つ読み進め、複合フィールドの終了要素に到達したときだけ完成した結果を返します。
            /// </summary>
            /// <param name="element">文書順に渡されるOOXML要素。</param>
            /// <param name="field">読み取りが完了した複合フィールド。途中の場合はnull。</param>
            /// <returns>この要素で複合フィールドの読み取りが完了した場合はtrue。</returns>
            public bool TryRead(
                OpenXmlElement element,
                [NotNullWhen(true)] out ComplexField? field)
            {
                if (element is Wordprocessing.FieldChar fieldCharacter)
                {
                    return TryReadFieldCharacter(fieldCharacter, out field);
                }

                if (element is Wordprocessing.FieldCode instructionElement)
                {
                    AppendInstruction(instructionElement);
                }
                else if (element is Wordprocessing.Text or
                    Wordprocessing.TabChar or
                    Wordprocessing.Break)
                {
                    AppendResultElement(element);
                }

                field = null;
                return false;
            }

            /// <summary>
            /// 複合フィールドの開始、結果開始、終了を読み取り状態へ反映します。
            /// </summary>
            /// <param name="fieldCharacter">状態を表すフィールド文字。</param>
            /// <param name="field">終了まで到達した場合の複合フィールド。それ以外はnull。</param>
            /// <returns>終了まで到達し、複合フィールドを構成できた場合はtrue。</returns>
            bool TryReadFieldCharacter(
                Wordprocessing.FieldChar fieldCharacter,
                [NotNullWhen(true)] out ComplexField? field)
            {
                var fieldCharacterType = fieldCharacter.FieldCharType?.Value;

                if (fieldCharacterType == Wordprocessing.FieldCharValues.Begin)
                {
                    StartField(fieldCharacter);
                }
                else if (fieldCharacterType == Wordprocessing.FieldCharValues.Separate)
                {
                    StartResult(fieldCharacter);
                }
                else if (fieldCharacterType == Wordprocessing.FieldCharValues.End)
                {
                    return TryCompleteField(fieldCharacter, out field);
                }

                field = null;
                return false;
            }

            /// <summary>
            /// 複合フィールドの結果開始前に現れた命令文字列を順番に連結します。
            /// </summary>
            /// <param name="instructionElement">命令文字列を保持するOOXML要素。</param>
            void AppendInstruction(Wordprocessing.FieldCode instructionElement)
            {
                if (fieldStart is not null && resultElements is null)
                {
                    fieldInstruction += instructionElement.Text;
                }
            }

            /// <summary>
            /// 複合フィールドの結果開始後に現れた表示用要素を保持します。
            /// </summary>
            /// <param name="element">表示値を構成する文字列、タブ、改行要素。</param>
            void AppendResultElement(OpenXmlElement element) =>
                resultElements?.Add(element);

            /// <summary>
            /// それまでの未完了状態を破棄し、新しい複合フィールドの読み取りを開始します。
            /// </summary>
            /// <param name="startElement">複合フィールドの開始要素。</param>
            void StartField(Wordprocessing.FieldChar startElement)
            {
                Reset();
                fieldStart = startElement;
            }

            /// <summary>
            /// 複合フィールドの命令部分を終了し、表示結果の収集を開始します。
            /// </summary>
            /// <param name="separatorElement">命令部分と表示結果を区切る要素。</param>
            void StartResult(Wordprocessing.FieldChar separatorElement)
            {
                if (fieldStart is not null)
                {
                    resultSeparator = separatorElement;
                    resultElements = [];
                }
            }

            /// <summary>
            /// 開始要素と表示結果が揃っている場合だけ複合フィールドを完成させ、読み取り状態を初期化します。
            /// </summary>
            /// <param name="fieldEnd">複合フィールドの終了要素。</param>
            /// <param name="field">完成した複合フィールド。必要な要素が不足している場合はnull。</param>
            /// <returns>複合フィールドを構成できた場合はtrue。</returns>
            bool TryCompleteField(
                Wordprocessing.FieldChar fieldEnd,
                [NotNullWhen(true)] out ComplexField? field)
            {
                if (fieldStart is null || resultSeparator is null || resultElements is null)
                {
                    field = null;
                    Reset();
                    return false;
                }

                field = new ComplexField(
                    fieldStart,
                    fieldInstruction,
                    resultSeparator,
                    fieldEnd,
                    resultElements);
                Reset();
                return true;
            }

            /// <summary>
            /// 次の複合フィールドを読み取れるよう、保持中のOOXML要素と文字列を破棄します。
            /// </summary>
            void Reset()
            {
                fieldStart = null;
                fieldInstruction = null;
                resultSeparator = null;
                resultElements = null;
            }
        }

        /// <summary>
        /// OOXML上で読み取りが完了した複合フィールドの構成要素を保持します。
        /// </summary>
        /// <param name="FieldStart">複合フィールドの開始要素。</param>
        /// <param name="Instruction">分割された要素を連結したフィールド命令。</param>
        /// <param name="ResultSeparator">命令部分と表示結果を区切る要素。</param>
        /// <param name="FieldEnd">複合フィールドの終了要素。</param>
        /// <param name="ResultElements">表示結果を構成する文字列、タブ、改行要素。</param>
        sealed record ComplexField(
            Wordprocessing.FieldChar FieldStart,
            string? Instruction,
            Wordprocessing.FieldChar ResultSeparator,
            Wordprocessing.FieldChar FieldEnd,
            IReadOnlyCollection<OpenXmlElement> ResultElements);
    }
}
