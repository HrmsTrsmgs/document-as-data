using System.Collections;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを取得するコレクションを表します。
/// </summary>
public class MergeFieldCollection : IEnumerable<MergeField>
{
    /// <summary>
    /// 列挙方法と生成済みMERGEFIELDをコレクションの生存期間中共有します。
    /// </summary>
    readonly MergeFieldReader reader;

    /// <summary>
    /// 文書の解析処理と生成済みMERGEFIELDのキャッシュを、コレクションの生存期間中共有できるようにします。
    /// </summary>
    /// <param name="document">MERGEFIELDを取得する文書。</param>
    internal MergeFieldCollection(Document document)
    {
        reader = new(document);
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
    sealed class MergeFieldReader
    {
        /// <summary>
        /// OOXML要素の取得元と、生成するMERGEFIELDの所属先です。
        /// </summary>
        readonly Document document;

        /// <summary>
        /// 同じOOXML要素からは、列挙方法にかかわらず同じMERGEFIELDを返すために保持します。
        /// </summary>
        readonly Dictionary<OpenXmlElement, MergeField> cache = [];

        /// <summary>
        /// 指定した文書に対する読み取り処理を作成します。
        /// </summary>
        /// <param name="document">読み取り対象の文書。</param>
        public MergeFieldReader(Document document)
        {
            this.document = document;
        }

        /// <summary>
        /// 文書要素を先頭から読み取り、認識できたMERGEFIELDを文書順に返します。
        /// 複合フィールドの読み取り状態は列挙ごとに作り直します。
        /// </summary>
        /// <returns>文書内で認識できたMERGEFIELD。</returns>
        public IEnumerable<MergeField> Read()
        {
            var complexFieldReader = new ComplexFieldReader();

            return
                from element in document.Elements
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
            if (TryParseMergeFieldName(field.Instruction?.Value, out var name))
            {
                mergeField =
                    GetOrCreateMergeField(
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
            if (TryParseMergeFieldName(field.Instruction, out var name))
            {
                mergeField =
                    GetOrCreateMergeField(
                        field.Element,
                        () => new(document, name, field.ResultTexts));
                return true;
            }

            mergeField = null;
            return false;
        }

        /// <summary>
        /// OOXML要素を識別子として生成済みのMERGEFIELDを再利用し、列挙と名前検索の同一性を保ちます。
        /// </summary>
        /// <param name="field">MERGEFIELDを表すOOXML要素。</param>
        /// <param name="create">未生成の場合にMERGEFIELDを作成する処理。</param>
        /// <returns>既存または新しく生成したMERGEFIELD。</returns>
        MergeField GetOrCreateMergeField(
            OpenXmlElement field,
            Func<MergeField> create)
        {
            if (cache.TryGetValue(field, out var mergeField))
            {
                return mergeField;
            }

            mergeField = create();
            cache.Add(field, mergeField);
            return mergeField;
        }

        /// <summary>
        /// フィールド命令がMERGEFIELDかを判定し、MERGEFIELD名を取り出します。
        /// </summary>
        /// <param name="fieldInstruction">OOXMLに保存されたフィールド命令。</param>
        /// <param name="name">MERGEFIELDだった場合の名前。それ以外の場合はnull。</param>
        /// <returns>MERGEFIELD命令だった場合はtrue。</returns>
        static bool TryParseMergeFieldName(
            string? fieldInstruction,
            [NotNullWhen(true)] out string? name)
        {
            const string fieldType = "MERGEFIELD";

            var instruction = fieldInstruction?.Trim();
            if (instruction is null ||
                !instruction.StartsWith(fieldType, StringComparison.OrdinalIgnoreCase))
            {
                name = null;
                return false;
            }

            name = instruction[fieldType.Length..].Trim().Trim('"');
            return true;
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
            Wordprocessing.FieldChar? begin;

            /// <summary>
            /// 開始要素から結果開始要素までに分割して保存された命令文字列を保持します。
            /// </summary>
            string? instruction;

            /// <summary>
            /// 結果開始後に現れた表示文字列を保持します。結果開始前はnullです。
            /// </summary>
            List<Wordprocessing.Text>? resultTexts;

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
                if (element is Wordprocessing.FieldChar fieldChar)
                {
                    return TryRead(fieldChar, out field);
                }

                if (element is Wordprocessing.FieldCode fieldCode)
                {
                    Read(fieldCode);
                }
                else if (element is Wordprocessing.Text text)
                {
                    Read(text);
                }

                field = null;
                return false;
            }

            /// <summary>
            /// 複合フィールドの開始、結果開始、終了を読み取り状態へ反映します。
            /// </summary>
            /// <param name="fieldChar">状態を表すフィールド文字。</param>
            /// <param name="field">終了まで到達した場合の複合フィールド。それ以外はnull。</param>
            /// <returns>終了まで到達し、複合フィールドを構成できた場合はtrue。</returns>
            bool TryRead(
                Wordprocessing.FieldChar fieldChar,
                [NotNullWhen(true)] out ComplexField? field)
            {
                var type = fieldChar.FieldCharType?.Value;

                if (type == Wordprocessing.FieldCharValues.Begin)
                {
                    StartField(fieldChar);
                }
                else if (type == Wordprocessing.FieldCharValues.Separate)
                {
                    StartResult();
                }
                else if (type == Wordprocessing.FieldCharValues.End)
                {
                    return TryCompleteField(out field);
                }

                field = null;
                return false;
            }

            /// <summary>
            /// 複合フィールドの結果開始前に現れた命令文字列を順番に連結します。
            /// </summary>
            /// <param name="fieldCode">命令文字列を保持するOOXML要素。</param>
            void Read(Wordprocessing.FieldCode fieldCode)
            {
                if (begin is not null && resultTexts is null)
                {
                    instruction += fieldCode.Text;
                }
            }

            /// <summary>
            /// 複合フィールドの結果開始後に現れた表示文字列を保持します。
            /// </summary>
            /// <param name="text">表示文字列を保持するOOXML要素。</param>
            void Read(Wordprocessing.Text text) =>
                resultTexts?.Add(text);

            /// <summary>
            /// それまでの未完了状態を破棄し、新しい複合フィールドの読み取りを開始します。
            /// </summary>
            /// <param name="fieldChar">複合フィールドの開始要素。</param>
            void StartField(Wordprocessing.FieldChar fieldChar)
            {
                begin = fieldChar;
                instruction = null;
                resultTexts = null;
            }

            /// <summary>
            /// 複合フィールドの命令部分を終了し、表示結果の収集を開始します。
            /// </summary>
            void StartResult()
            {
                if (begin is not null)
                {
                    resultTexts = [];
                }
            }

            /// <summary>
            /// 開始要素と表示結果が揃っている場合だけ複合フィールドを完成させ、読み取り状態を初期化します。
            /// </summary>
            /// <param name="field">完成した複合フィールド。必要な要素が不足している場合はnull。</param>
            /// <returns>複合フィールドを構成できた場合はtrue。</returns>
            bool TryCompleteField(
                [NotNullWhen(true)] out ComplexField? field)
            {
                if (begin is null || resultTexts is null)
                {
                    field = null;
                    Reset();
                    return false;
                }

                field = new ComplexField(begin, instruction, resultTexts);
                Reset();
                return true;
            }

            /// <summary>
            /// 次の複合フィールドを読み取れるよう、保持中のOOXML要素と文字列を破棄します。
            /// </summary>
            void Reset()
            {
                begin = null;
                instruction = null;
                resultTexts = null;
            }
        }

        /// <summary>
        /// OOXML上で読み取りが完了した複合フィールドの構成要素を保持します。
        /// </summary>
        /// <param name="Element">キャッシュの識別子となる開始要素。</param>
        /// <param name="Instruction">分割された要素を連結したフィールド命令。</param>
        /// <param name="ResultTexts">表示結果を構成する文字列要素。</param>
        sealed record ComplexField(
            OpenXmlElement Element,
            string? Instruction,
            IReadOnlyCollection<Wordprocessing.Text> ResultTexts);
    }
}
