using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// 生成ソースの構文木から宣言やコメントを取り出す低レベルなテスト補助です。
/// </summary>
static class GeneratedSourceAssertions
{
    extension(IEnumerable<string> self)
    {
        /// <summary>
        /// 生成ソースに含まれる型宣言を取得します。
        /// </summary>
        /// <returns>ソース内の型宣言。</returns>
        internal IEnumerable<TypeDeclarationSyntax> TypeDeclarations() =>
            from source in self
            from type in
                CSharpSyntaxTree
                    .ParseText(source)
                    .GetCompilationUnitRoot()
                    .DescendantNodes()
                    .OfType<TypeDeclarationSyntax>()
            select type;

        /// <summary>
        /// 生成ソースから指定した名前の型宣言を取得します。
        /// </summary>
        /// <param name="name">取得する型名。</param>
        /// <returns>指定した型宣言。</returns>
        internal TypeDeclarationSyntax TypeDeclaration(string name) =>
            (
                from type in self.TypeDeclarations()
                where type.Identifier.ValueText == name
                select type
            ).Single();
    }

    extension(MemberDeclarationSyntax self)
    {
        /// <summary>
        /// 宣言に付けられたXMLコメントの summary 本文を取得します。
        /// </summary>
        /// <returns>summary 本文。summary がない場合は null。</returns>
        internal string? SummaryText()
        {
            var trivia = self.GetLeadingTrivia()
                .Select(it => it.GetStructure())
                .OfType<DocumentationCommentTriviaSyntax>()
                .SingleOrDefault();

            var summary = trivia
                ?.Content
                .OfType<XmlElementSyntax>()
                .SingleOrDefault(
                    it => it.StartTag.Name.LocalName.ValueText == "summary");

            return summary == null
                ? null
                : SingleLineText(summary);
        }
    }

    /// <summary>
    /// XMLコメントのテキストを一行へ整形します。
    /// エンティティの前後も別トークンになるため、トークンは空白を挟まず連結し、
    /// 行の境界だけを空白で区切ります。
    /// </summary>
    /// <param name="element">本文を取得するXMLコメント要素。</param>
    /// <returns>空行と行頭・行末の空白を除いて一行にまとめた本文。</returns>
    static string SingleLineText(XmlElementSyntax element) =>
        string.Join(
            " ",
            from line in
                string.Concat(
                    element.Content
                        .OfType<XmlTextSyntax>()
                        .SelectMany(it => it.TextTokens)
                        .Select(it => it.ValueText))
                    .Split('\n')
            let text = line.Trim()
            where text != ""
            select text);

    extension(TypeDeclarationSyntax self)
    {
        /// <summary>
        /// 指定した型宣言から指定した名前のプロパティ宣言を取得します。
        /// </summary>
        /// <param name="propertyName">取得するプロパティ名。</param>
        /// <returns>指定したプロパティ宣言。</returns>
        internal PropertyDeclarationSyntax PropertyDeclaration(
            string propertyName) =>
            (
                from property in self.Members.OfType<PropertyDeclarationSyntax>()
                where property.Identifier.ValueText == propertyName
                select property
            ).Single();

        /// <summary>
        /// 指定した型宣言から、指定した名前の引数なしメソッド宣言を取得します。
        /// </summary>
        /// <param name="methodName">取得するメソッド名。</param>
        /// <returns>指定した引数なしメソッド宣言。</returns>
        internal MethodDeclarationSyntax MethodDeclaration(string methodName) =>
            (
                from method in self.Members.OfType<MethodDeclarationSyntax>()
                where method.Identifier.ValueText == methodName
                where method.ParameterList.Parameters.Count == 0
                select method
            ).Single();

        /// <summary>
        /// 指定した型宣言から、名前と引数型が一致するメソッド宣言を取得します。
        /// </summary>
        /// <param name="methodName">取得するメソッド名。</param>
        /// <param name="parameterTypeName">唯一の引数に指定された型名。</param>
        /// <returns>指定したメソッド宣言。</returns>
        internal MethodDeclarationSyntax MethodDeclaration(
            string methodName,
            string parameterTypeName) =>
            (
                from method in self.Members.OfType<MethodDeclarationSyntax>()
                where method.Identifier.ValueText == methodName
                where method.ParameterList.Parameters.Count == 1
                where method.ParameterList.Parameters[0].Type?.ToString()
                    == parameterTypeName
                select method
            ).Single();
    }
}
