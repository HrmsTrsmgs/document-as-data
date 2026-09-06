using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内の文字列Content Controlを表します。
/// </summary>
public class TextContentControl : ContentControl, DocumentTextItem
{
    /// <summary>
    /// 文書内のOOXML要素への参照を保持し、Textの読み書きを同じ要素へ反映できるようにします。
    /// </summary>
    /// <param name="document">Content Controlが属する文書。</param>
    /// <param name="element">Content Controlを構成するOOXML要素。</param>
    internal TextContentControl(
        Document document,
        Wordprocessing.SdtElement element)
        : base(document, element)
    {
    }

    /// <summary>
    /// Content Controlの文字列を取得または設定します。
    /// </summary>
    /// <remarks>
    /// プレースホルダー表示中は、表示用文字列ではなく空文字列を返します。
    /// 値を設定すると、プレースホルダー表示状態を解除します。
    /// </remarks>
    /// <exception cref="InvalidOperationException">
    /// 値の設定時に文字列要素が存在しない場合。
    /// </exception>
    public string Text
    {
        get => IsShowingPlaceholder
            ? ""
            : WordTextValue.Read(Element.Descendants());
        set
        {
            var texts = Element.Descendants<Wordprocessing.Text>().ToArray();

            RemovePlaceholderState();
            RemoveControlCharacters();
            ReplaceFirstText(texts.First(), value);
            foreach (var text in texts[1..])
            {
                text.Text = "";
            }
        }
    }

    /// <summary>
    /// 前回の値を構成していたタブと改行が、再設定した値の前に残らないよう除去します。
    /// </summary>
    void RemoveControlCharacters()
    {
        // OOXML要素を列挙しながら削除すると次の要素をたどれないため、削除対象を先に確定します。
        foreach (var valueElement in
            (
                from valueElement in Element.Descendants()
                where valueElement is Wordprocessing.TabChar or Wordprocessing.Break
                select valueElement
            ).ToArray())
        {
            valueElement.Remove();
        }
    }

    /// <summary>
    /// 最初の文字列要素を、Wordが表示できる文字列、タブ、改行要素へ置き換えます。
    /// 空文字列では次回の設定先を失わないよう、文字列要素自体を残します。
    /// </summary>
    /// <param name="text">置換対象の最初の文字列要素。</param>
    /// <param name="value">設定する値。</param>
    static void ReplaceFirstText(Wordprocessing.Text text, string value)
    {
        if (value.IsEmpty())
        {
            text.Text = "";
            return;
        }

        var parent = text.Parent ?? throw new InvalidOperationException();

        text.Remove();
        parent.Append(WordTextValue.CreateElements(value));
    }

    /// <summary>
    /// このContent Controlがプレースホルダー表示中かを取得します。
    /// </summary>
    bool IsShowingPlaceholder =>
        Element.PropertyElements<Wordprocessing.ShowingPlaceholder>().Any();

    /// <summary>
    /// 値の書き込み後に表示文字列がプレースホルダーとして扱われないよう、表示状態を解除します。
    /// </summary>
    void RemovePlaceholderState() =>
        Element.PropertyElements<Wordprocessing.ShowingPlaceholder>()
            .SingleOrDefault()?.Remove();
}
