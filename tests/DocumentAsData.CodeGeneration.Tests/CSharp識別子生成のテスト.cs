using FluentAssertions;

using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class CSharp識別子生成のテスト
{
    const string キャメルケースIdentifierDocumentFilePath =
        @"TestData\コード生成\salesReport.docx";
    const string キャメルケースMergeFieldDocumentFilePath =
        @"TestData\コード生成\customerData.docx";
    const string キャメルケースTextContentControlDocumentFilePath =
        @"TestData\コード生成\customerForm.docx";
    const string キャメルケースCheckBoxDocumentFilePath =
        @"TestData\コード生成\acceptanceForm.docx";
    const string キャメルケースDatePickerDocumentFilePath =
        @"TestData\コード生成\deliveryForm.docx";

    [Theory]
    [InlineData("salesReport", "SalesReport")]
    [InlineData("salesData", "SalesData")]
    [InlineData("salesDetail", "SalesDetail")]
    [InlineData("customerId", "CustomerId")]
    public void キャメルケースの文書由来名はPascalCase識別子へ変換します(
        string sourceName,
        string identifierBody)
    {
        CSharpIdentifier
            .ToCSharpIdentifier(sourceName)
            .Should().Be(identifierBody);
    }

    [Fact]
    public void キャメルケース文書ファイル名はPascalCaseのDocument型名へ変換します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierDocumentFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain("SalesReportDocument");
    }

    [Fact]
    public void キャメルケースMERGEFIELD名はPascalCaseのプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースMergeFieldDocumentFilePath))
            .GeneratedType("CustomerDataDocument")
            .GetProperty("CustomerName")
            .Should().NotBeNull();
    }

    [Fact]
    public void キャメルケース文字列ContentControlのTagはPascalCaseのプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースTextContentControlDocumentFilePath))
            .GeneratedType("CustomerFormDocument")
            .GetProperty("CustomerName")
            .Should().NotBeNull();
    }

    [Fact]
    public void キャメルケースCheckBoxのTagはPascalCaseのプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースCheckBoxDocumentFilePath))
            .GeneratedType("AcceptanceFormDocument")
            .GetProperty("TermsAccepted")
            .Should().NotBeNull();
    }

    [Fact]
    public void キャメルケースDatePickerのTagはPascalCaseのプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースDatePickerDocumentFilePath))
            .GeneratedType("DeliveryFormDocument")
            .GetProperty("DeliveryDate")
            .Should().NotBeNull();
    }

    [Theory]
    [InlineData("sales_detail", "SalesDetail")]
    [InlineData("sales-detail", "SalesDetail")]
    [InlineData("sales detail", "SalesDetail")]
    public void ASCII名は区切り文字を単語境界としてPascalCase識別子へ変換します(
        string sourceName,
        string identifierBody)
    {
        CSharpIdentifier
            .ToCSharpIdentifier(sourceName)
            .Should().Be(identifierBody);
    }

    [Fact]
    public void ASCII名は全大文字の単語をPascalCase識別子へ正規化します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("SALES_DETAIL1")
            .Should().Be("SalesDetail1");
    }

    [Fact]
    public void ASCII名は二文字頭字語を両方大文字の識別子へ変換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("IO_stream")
            .Should().Be("IOStream");
    }

    [Fact]
    public void ASCII名はIdを二文字頭字語の例外として変換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("customer_ID")
            .Should().Be("CustomerId");
    }

    [Theory]
    [InlineData("商品_明細", "商品_明細")]
    [InlineData("商品-明細", "商品_明細")]
    [InlineData("sales商品-detail", "Sales商品_detail")]
    public void 非ASCIIを含む名は内部の単語境界を推測しません(
        string sourceName,
        string identifierBody)
    {
        CSharpIdentifier
            .ToCSharpIdentifier(sourceName)
            .Should().Be(identifierBody);
    }

    [Theory]
    [InlineData('!')]
    [InlineData('"')]
    [InlineData('#')]
    [InlineData('%')]
    [InlineData('&')]
    [InlineData('\'')]
    [InlineData('*')]
    [InlineData(',')]
    [InlineData('.')]
    [InlineData('/')]
    [InlineData(':')]
    [InlineData(';')]
    [InlineData('?')]
    [InlineData('@')]
    [InlineData('\\')]
    public void OtherPunctuationに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('(')]
    [InlineData('[')]
    [InlineData('{')]
    public void OpenPunctuationに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void ClosePunctuationに分類される右丸かっこはアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price)rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void ClosePunctuationに分類される右角かっこはアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price]rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void ClosePunctuationに分類される右波かっこはアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('+')]
    [InlineData('<')]
    [InlineData('=')]
    [InlineData('>')]
    [InlineData('|')]
    [InlineData('~')]
    public void MathSymbolに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('^')]
    [InlineData('`')]
    public void ModifierSymbolに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void CurrencySymbolに分類されるASCII記号はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price$rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('！')]
    [InlineData('＂')]
    [InlineData('＃')]
    [InlineData('％')]
    [InlineData('＆')]
    [InlineData('＇')]
    [InlineData('＊')]
    [InlineData('，')]
    [InlineData('．')]
    [InlineData('／')]
    [InlineData('：')]
    [InlineData('；')]
    [InlineData('？')]
    [InlineData('＠')]
    [InlineData('＼')]
    public void OtherPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('（')]
    [InlineData('［')]
    [InlineData('｛')]
    public void OpenPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('）')]
    [InlineData('］')]
    [InlineData('｝')]
    public void ClosePunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void DashPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price－rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('＋')]
    [InlineData('＜')]
    [InlineData('＝')]
    [InlineData('＞')]
    [InlineData('｜')]
    [InlineData('～')]
    public void MathSymbolに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData('＾')]
    [InlineData('｀')]
    public void ModifierSymbolに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        CSharpIdentifier
            .ToCSharpIdentifier($"price{character}rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void CurrencySymbolに分類される全角ASCII相当記号はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price＄rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void OtherSymbolに分類される漢字構成記述文字はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("商品⿰明細")
            .Should().Be("商品_明細");
    }

    [Fact]
    public void SpaceSeparatorに分類される全角空白はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price\u3000rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void Controlに分類される制御文字はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price\u0001rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void OtherNumberに分類される数値文字はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price\u00B2rate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void EnclosingMarkに分類される囲み結合記号はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price\u20DDrate")
            .Should().Be("Price_rate");
    }

    [Fact]
    public void PrivateUseに分類される私用領域文字はアンダースコアへ置換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("price\uE000rate")
            .Should().Be("Price_rate");
    }

    [Theory]
    [InlineData("2026_sales", "_2026Sales")]
    [InlineData("2026商品", "_2026商品")]
    public void 数字から始まる名は有効なCSharp識別子へ補正します(
        string sourceName,
        string identifierBody)
    {
        CSharpIdentifier
            .ToCSharpIdentifier(sourceName)
            .Should().Be(identifierBody);
    }

    [Theory]
    [InlineData("\u203Fsales", "_\u203Fsales")]
    [InlineData("\u0301sales", "_\u0301sales")]
    [InlineData("\u0903sales", "_\u0903sales")]
    [InlineData("\u200Csales", "_\u200Csales")]
    public void 開始文字として使用できないカテゴリで始まる名は有効なCSharp識別子へ補正します(
        string sourceName,
        string identifierBody)
    {
        CSharpIdentifier
            .ToCSharpIdentifier(sourceName)
            .Should().Be(identifierBody);
    }

    [Fact]
    public void CSharpキーワードと同じ文書由来名はキーワードでない識別子へ変換します()
    {
        CSharpIdentifier
            .ToCSharpIdentifier("class")
            .Should().Be("Class");
    }
}
