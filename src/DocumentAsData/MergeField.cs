using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを表します。
/// </summary>
public class MergeField
{
    /// <summary>
    /// 単純形式または複合形式のOOXML要素への値アクセスを保持します。
    /// </summary>
    readonly MergeFieldContent content;

    /// <summary>
    /// 単純フィールド要素自体を値の読み書き対象とするMERGEFIELDを作成します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="field">MERGEFIELDを構成する単純フィールド要素。</param>
    /// <param name="name">MERGEFIELDの名前。</param>
    internal MergeField(
        Document document,
        Wordprocessing.SimpleField field,
        string name)
    {
        Document = document;
        content = new SimpleFieldContent(field);
        Name = name;
    }

    /// <summary>
    /// 開始要素と結果文字列が分かれた複合形式のMERGEFIELDを作成します。
    /// </summary>
    /// <param name="document">MERGEFIELDが属する文書。</param>
    /// <param name="name">MERGEFIELDの名前。</param>
    /// <param name="fieldStart">複合フィールドの開始要素。</param>
    /// <param name="resultSeparator">命令部分と表示結果を区切る要素。</param>
    /// <param name="fieldEnd">複合フィールドの終了要素。</param>
    /// <param name="valueElements">MERGEFIELDの表示値を構成する文字列、タブ、改行要素。</param>
    internal MergeField(
        Document document,
        string name,
        Wordprocessing.FieldChar fieldStart,
        Wordprocessing.FieldChar resultSeparator,
        Wordprocessing.FieldChar fieldEnd,
        IEnumerable<OpenXmlElement> valueElements)
    {
        Document = document;
        content = new ComplexFieldContent(
            fieldStart,
            resultSeparator,
            fieldEnd,
            valueElements);
        Name = name;
    }

    /// <summary>
    /// MERGEFIELDが属する文書を取得します。
    /// </summary>
    public Document Document { get; }

    /// <summary>
    /// MERGEFIELDの名前を取得します。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// MERGEFIELDの値を取得または設定します。
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// 値の設定時に表示結果が存在しないか、複合フィールドの結果領域を取得できない場合。
    /// </exception>
    public string Value
    {
        get => content.Value;
        set => content.Value = value;
    }

    /// <summary>
    /// MERGEFIELDのOOXML形式に依存した値の読み書きを表します。
    /// </summary>
    abstract class MergeFieldContent
    {
        /// <summary>
        /// OOXMLの表示値を取得または設定します。
        /// </summary>
        internal abstract string Value { get; set; }
    }

    /// <summary>
    /// 単純形式のMERGEFIELDの値を読み書きします。
    /// </summary>
    sealed class SimpleFieldContent : MergeFieldContent
    {
        /// <summary>
        /// MERGEFIELDを構成する単純フィールド要素です。
        /// </summary>
        readonly Wordprocessing.SimpleField simpleField;

        /// <summary>
        /// 指定した単純フィールド要素の値を読み書きする内部表現を作成します。
        /// </summary>
        /// <param name="field">MERGEFIELDを構成する単純フィールド要素。</param>
        internal SimpleFieldContent(Wordprocessing.SimpleField field)
        {
            simpleField = field;
        }

        /// <inheritdoc />
        internal override string Value
        {
            get => WordTextValue.Read(simpleField.Descendants());
            set
            {
                simpleField.RemoveAllChildren();
                simpleField.AppendChild(
                    new Wordprocessing.Run(WordTextValue.CreateElements(value)));
                // 書き換えた表示結果をWordが古い結果として扱わないようにします。
                simpleField.Dirty = null;
            }
        }
    }

    /// <summary>
    /// 複合形式のMERGEFIELDの値を読み書きします。
    /// </summary>
    sealed class ComplexFieldContent : MergeFieldContent
    {
        /// <summary>
        /// 複合フィールドの開始要素です。
        /// 値の書き込み後に、表示結果が古いことを示す状態を解除するため保持します。
        /// </summary>
        readonly Wordprocessing.FieldChar fieldStart;

        /// <summary>
        /// 複合フィールドの命令部分と表示結果を区切る要素です。
        /// </summary>
        readonly Wordprocessing.FieldChar resultSeparator;

        /// <summary>
        /// 複合フィールドの終了要素です。
        /// </summary>
        readonly Wordprocessing.FieldChar fieldEnd;

        /// <summary>
        /// 表示値を構成する文字列、タブ、改行要素を文書順に保持します。
        /// 読み取り時はこれらを連結し、書き込み時は置換後の要素へ更新します。
        /// </summary>
        readonly List<OpenXmlElement> valueElements;

        /// <summary>
        /// 複合フィールドの各構成要素から、値を読み書きする内部表現を作成します。
        /// </summary>
        /// <param name="fieldStart">複合フィールドの開始要素。</param>
        /// <param name="resultSeparator">命令部分と表示結果を区切る要素。</param>
        /// <param name="fieldEnd">複合フィールドの終了要素。</param>
        /// <param name="valueElements">表示値を構成する文字列、タブ、改行要素。</param>
        internal ComplexFieldContent(
            Wordprocessing.FieldChar fieldStart,
            Wordprocessing.FieldChar resultSeparator,
            Wordprocessing.FieldChar fieldEnd,
            IEnumerable<OpenXmlElement> valueElements)
        {
            this.fieldStart = fieldStart;
            this.resultSeparator = resultSeparator;
            this.fieldEnd = fieldEnd;
            this.valueElements = valueElements.ToList();
        }

        /// <inheritdoc />
        internal override string Value
        {
            get => WordTextValue.Read(valueElements);
            set => SetValue(value);
        }

        /// <summary>
        /// 複合フィールドの古い表示結果を除去し、新しい値を持つ結果だけに置き換えます。
        /// 空文字列でも、次回の書き込み位置として空の文字列要素を結果領域に残します。
        /// </summary>
        /// <param name="value">設定する値。</param>
        /// <exception cref="InvalidOperationException">
        /// 表示結果が存在しないか、結果領域の境界を取得できない場合。
        /// </exception>
        void SetValue(string value)
        {
            if (valueElements.Count == 0)
            {
                throw new InvalidOperationException();
            }

            var resultStart = resultSeparator.Parent ??
                throw new InvalidOperationException();
            var resultEnd = fieldEnd.Parent ??
                throw new InvalidOperationException();
            if (resultStart.Parent != resultEnd.Parent)
            {
                throw new InvalidOperationException();
            }

            // OOXML要素を列挙しながら削除すると次の兄弟をたどれないため、削除対象を先に確定します。
            foreach (var oldElement in
                resultStart.ElementsAfter()
                    .TakeWhile(it => it != resultEnd)
                    .ToArray())
            {
                oldElement.Remove();
            }

            var newValueElements = WordTextValue.CreateElements(value)
                .DefaultIfEmpty(new Wordprocessing.Text())
                .ToArray();
            resultEnd.InsertBeforeSelf(
                new Wordprocessing.Run(newValueElements));
            valueElements.Clear();
            valueElements.AddRange(newValueElements);
            // 書き換えた表示結果をWordが古い結果として扱わないようにします。
            fieldStart.Dirty = null;
        }
    }
}
