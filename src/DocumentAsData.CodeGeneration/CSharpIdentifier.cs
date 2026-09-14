using System.Globalization;

namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// 文書由来の名前をC#識別子へ変換します。
/// </summary>
static class CSharpIdentifier
{
    extension(string self)
    {
        /// <summary>
        /// C#の識別子比較で無視される先頭の@と書式文字を除き、衝突診断に使う名前を取得します。
        /// 生成コードや文書内の元名自体は変更しません。
        /// </summary>
        internal string IdentifierComparisonKey =>
            string.Concat(
                from character in self.StartsWith('@') ? self[1..] : self
                where char.GetUnicodeCategory(character) != UnicodeCategory.Format
                select character);
    }

    /// <summary>
    /// 文書内で使用されている名前を、生成コードで使用するC#識別子へ変換します。
    /// </summary>
    /// <param name="sourceName">文書内で使用されている名前。</param>
    /// <returns>C#識別子として使用できる名前。</returns>
    internal static string ToCSharpIdentifier(this string sourceName) =>
        EnsureValidIdentifierStart(
            ContainsNonAscii(sourceName)
                ? CapitalizeFirstLetter(
                    ReplaceInvalidIdentifierPartCharacters(
                        sourceName.Replace('-', '_').Replace(' ', '_')))
                : AsciiIdentifier(sourceName));

    /// <summary>
    /// ASCIIだけで構成された名前を、区切り文字と大文字小文字からPascalCaseへ変換します。
    /// </summary>
    static string AsciiIdentifier(string sourceName) =>
        string.Concat(
            from word in sourceName.Split(['_', '-', ' '])
            select PascalCaseWord(ReplaceInvalidIdentifierPartCharacters(word)));

    /// <summary>
    /// C#識別子の開始文字として使えない場合に、先頭へアンダースコアを補います。
    /// </summary>
    static string EnsureValidIdentifierStart(string identifier) =>
        identifier.Length == 0 || IsIdentifierStartCharacter(identifier[0])
            ? identifier
            : $"_{identifier}";

    /// <summary>
    /// C#識別子の開始文字として使用できるUnicodeカテゴリかどうかを判定します。
    /// </summary>
    static bool IsIdentifierStartCharacter(char character) =>
        character == '_'
            || char.GetUnicodeCategory(character) is
                UnicodeCategory.UppercaseLetter
                or UnicodeCategory.LowercaseLetter
                or UnicodeCategory.TitlecaseLetter
                or UnicodeCategory.ModifierLetter
                or UnicodeCategory.OtherLetter
                or UnicodeCategory.LetterNumber;

    /// <summary>
    /// C#識別子の構成文字として使用できない文字をアンダースコアへ置換します。
    /// </summary>
    static string ReplaceInvalidIdentifierPartCharacters(string sourceName) =>
        string.Concat(
            from character in sourceName
            select IsIdentifierPartCharacter(character)
                ? character
                : '_');

    /// <summary>
    /// C#識別子の2文字目以降で使用できるUnicodeカテゴリかどうかを判定します。
    /// </summary>
    static bool IsIdentifierPartCharacter(char character) =>
        char.GetUnicodeCategory(character) is
            UnicodeCategory.UppercaseLetter
            or UnicodeCategory.LowercaseLetter
            or UnicodeCategory.TitlecaseLetter
            or UnicodeCategory.ModifierLetter
            or UnicodeCategory.OtherLetter
            or UnicodeCategory.LetterNumber
            or UnicodeCategory.DecimalDigitNumber
            or UnicodeCategory.ConnectorPunctuation
            or UnicodeCategory.NonSpacingMark
            or UnicodeCategory.SpacingCombiningMark
            or UnicodeCategory.Format;

    /// <summary>
    /// ASCII以外の文字を含むかどうかを判定します。
    /// </summary>
    static bool ContainsNonAscii(string sourceName) =>
        sourceName.Any(it => !char.IsAscii(it));

    /// <summary>
    /// ASCII名から切り出した1単語をPascalCaseの構成要素へ変換します。
    /// </summary>
    static string PascalCaseWord(string word) =>
        CapitalizeFirstLetter(
            ShouldNormalizeUpperCaseWord(word)
                ? word.ToLowerInvariant()
                : word);

    /// <summary>
    /// 全大文字単語を通常のPascalCase単語へ正規化するかどうかを判定します。
    /// </summary>
    static bool ShouldNormalizeUpperCaseWord(string word) =>
        IsUpperCaseWord(word)
            && !IsPreservedTwoLetterAcronym(word);

    /// <summary>
    /// 単語内の英字がすべて大文字かどうかを判定します。
    /// </summary>
    static bool IsUpperCaseWord(string word) =>
        word.Any(char.IsLetter)
            && word
                .Where(char.IsLetter)
                .All(char.IsUpper);

    /// <summary>
    /// .NET命名規則に合わせ、両方大文字のまま残す2文字頭字語かどうかを判定します。
    /// </summary>
    static bool IsPreservedTwoLetterAcronym(string word) =>
        word != "ID"
            && word.Count(char.IsLetter) == 2;

    /// <summary>
    /// 先頭文字だけを大文字化します。
    /// </summary>
    static string CapitalizeFirstLetter(string word) =>
        word.Length == 0
            ? word
            : $"{char.ToUpperInvariant(word[0])}{word[1..]}";
}
