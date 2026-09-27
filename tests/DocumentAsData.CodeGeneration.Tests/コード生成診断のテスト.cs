using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string MergeFieldNameCollisionDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD名衝突.docx";
    const string DifferentItemTypeNameCollisionDocumentFilePath =
        @"TestData\コード生成\異種項目名衝突.docx";
    const string InvalidDocumentNameFilePath =
        @"TestData\コード生成\---.docx";
    const string CheckBoxAndDatePickerNameCollisionDocumentFilePath =
        @"TestData\コード生成\CheckBoxとDatePickerの名前衝突.docx";

    [Fact]
    public void 自動変換後に同じMERGEFIELDプロパティ名となる場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(MergeFieldNameCollisionDocumentFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "CustomerId",
                    ["customer_id", "customer-id"]));
    }

    [Fact]
    public void MERGEFIELDプロパティ名の衝突を自動的な連番追加では解消しません()
    {
        GeneratedCodeInspection
            .GenerateSources(MergeFieldNameCollisionDocumentFilePath)
            .Should().NotContain(
                source => source.Contains(
                    "CustomerId2",
                    StringComparison.Ordinal));
    }

    [Fact]
    public void NameMappingsで生成名を変更するとMERGEFIELDプロパティ名の衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options =>
                    options.NameMappings["customer-id"] = "CustomerIdDash")
            .Should().BeEmpty();
    }

    [Fact]
    public void 生成プロパティ名が生成Document型のReadメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Read")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Read",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型のOpenメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Open")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Open",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型のReplaceメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Replace")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Replace",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのSaveメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Save")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "Save", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのSaveAsメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "SaveAs")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SaveAs", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのCloseメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Close")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "Close", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのDisposeメソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "Dispose")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "Dispose", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのMergeFieldsコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "MergeFields")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "MergeFields",
                    ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのContentControlsコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\文字列ContentControl.docx",
                options => options.NameMappings["CustomerName"] = "ContentControls")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "ContentControls",
                    ["CustomerName"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのCheckBoxesコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\チェック済みCheckBox.docx",
                options => options.NameMappings["Agreement"] = "CheckBoxes")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CheckBoxes", ["Agreement"]));
    }

    [Fact]
    public void 生成プロパティ名がDocumentのDatePickersコレクションと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\日付選択ContentControl.docx",
                options => options.NameMappings["DeliveryDate"] = "DatePickers")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "DatePickers", ["DeliveryDate"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型の項目検証メソッドと衝突した場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "ValidateRequiredItems")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "ValidateRequiredItems", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Document型名と同じ場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "CustomerDataDocument")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CustomerDataDocument", ["customerName"]));
    }

    [Fact]
    public void 生成プロパティ名が生成Data型名と同じ場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "CustomerDataData")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "CustomerDataData", ["customerName"]));
    }

    [Theory]
    [InlineData("OrderDocument")]
    [InlineData("OrderData")]
    public void NameMappingsで変更した生成型名とプロパティ名が同じ場合にもエラーを診断します(string propertyName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings = new()
                {
                    ["customerData"] = "Order",
                    ["customerName"] = propertyName
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, propertyName, ["customerName"]));
    }

    [Fact]
    public void 同じ生成Document型内の異なる種類のプロパティ名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(DifferentItemTypeNameCollisionDocumentFilePath)
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "CustomerId"
                && it.SourceNames.Contains("customer_id")
                && it.SourceNames.Contains("customer-id"));
    }

    [Fact]
    public void 異なる明細に同じTagがあっても生成名の衝突と診断しません()
    {
        // Itemsの二つの明細には、それぞれItemNameが一つずつあります。
        GeneratedCodeInspection
            .GenerateDiagnostics(@"TestData\コード生成\repeatingItems.docx")
            .Should().BeEmpty();
    }

    [Fact]
    public void 異なる明細の同名MERGEFIELDを生成名の衝突と診断しません()
    {
        // 二つの明細には、それぞれCodeというMERGEFIELDが一つずつあります。
        GeneratedCodeInspection
            .GenerateDiagnostics(@"TestData\コード生成\repeatingKinds.docx")
            .Should().NotContain(it => it.GeneratedName == "Code");
    }

    [Fact]
    public void 異なる明細の同名CheckBoxを生成名の衝突と診断しません()
    {
        // 二つの明細には、それぞれAgreementというCheckBoxが一つずつあります。
        GeneratedCodeInspection
            .GenerateDiagnostics(@"TestData\コード生成\repeatingKinds.docx")
            .Should().NotContain(it => it.GeneratedName == "Agreement");
    }

    [Fact]
    public void 異なる明細の同名DatePickerを生成名の衝突と診断しません()
    {
        // 二つの明細には、それぞれDeliveryDateというDatePickerが一つずつあります。
        GeneratedCodeInspection
            .GenerateDiagnostics(@"TestData\コード生成\repeatingKinds.docx")
            .Should().NotContain(it => it.GeneratedName == "DeliveryDate");
    }

    [Fact]
    public void 通常項目と明細内項目が同名でも生成名の衝突と診断しません()
    {
        // 外側のItemNameと、Items内の各明細のItemNameは別の生成データ型に属します。
        GeneratedCodeInspection
            .GenerateDiagnostics(@"TestData\コード生成\repeatingOutside.docx")
            .Should().BeEmpty();
    }

    [Fact(Skip = "RS-02: 生成セクション名と既存メンバー名の衝突診断をレビューしてからRedにします。")]
    public void 繰り返しセクションの生成プロパティ名が既存メンバー名と衝突すると診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\repeatingTemplate.docx",
                options => options.NameMappings["Items"] = "RepeatingSections")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "RepeatingSections", ["Items"]));
    }

    [Fact(Skip = "RS-02: セクションと通常項目の生成名が同じ場合の診断をレビューしてからRedにします。")]
    public void 繰り返しセクションと通常項目の生成プロパティ名が衝突すると診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\repeatingOutside.docx",
                options => options.NameMappings["Items"] = "ItemName")
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "ItemName"
                && it.SourceNames.Contains("Items")
                && it.SourceNames.Contains("ItemName"));
    }

    [Fact(Skip = "RS-02: 二つのセクションが同じ明細データ型名を生成する場合の診断をレビューしてからRedにします。")]
    public void 異なるセクションの明細データ型名が衝突すると診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\repeatingSections.docx",
                options => options.NameMappings["Options"] = "Items")
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "ItemsData"
                && it.SourceNames.Contains("Items")
                && it.SourceNames.Contains("Options"));
    }

    [Fact]
    public void 同じ生成名に三つ以上の文書項目が対応した場合にすべての元名を一つの診断へ含めます()
    {
        // 固定文書にはCustomerName・Address・Telephoneの単純MERGEFIELDが一つずつあります。
        // 元の名前は重複しておらず、生成名の割り当てによって初めて3項目が衝突します。
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\3項目のMERGEFIELD.docx",
                options => options.NameMappings = new()
                {
                    ["CustomerName"] = "SharedValue",
                    ["Address"] = "SharedValue",
                    ["Telephone"] = "SharedValue"
                })
            .Should().ContainSingle()
            .Which.Should().BeEquivalentTo(
                new CodeGenerationDiagnostic(
                    true,
                    "SharedValue",
                    ["CustomerName", "Address", "Telephone"]));
    }

    [Fact]
    public void 区切り文字だけの文書名はCSharp識別子を生成できないため診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(InvalidDocumentNameFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "",
                    ["---"],
                    "---"));
    }

    [Fact]
    public void CheckBoxとDatePickerの生成プロパティ名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                CheckBoxAndDatePickerNameCollisionDocumentFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "DataItem",
                    ["data_item", "data-item"]));
    }

    [Fact]
    public void NameMappingsで生成プロパティ名を空にした場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = "")
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "", ["customerName"]));
    }

    [Fact]
    public void 書式文字だけが異なる生成プロパティ名も同じ識別子として診断します()
    {
        // U+200CはC#の識別子の比較では無視される書式文字です。
        // 文字列としては異なる名前でも、生成した二つのプロパティは衝突します。
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options => options.NameMappings = new()
                {
                    ["customer_id"] = "Shared\u200CValue",
                    ["customer-id"] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", ["customer_id", "customer-id"]));
    }

    [Fact]
    public void 先頭のアットマークだけが異なる生成プロパティ名も同じ識別子として診断します()
    {
        // C#の識別子の先頭に付ける@は、識別子名自体には含まれません。
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options => options.NameMappings = new()
                {
                    ["customer_id"] = "@SharedValue",
                    ["customer-id"] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", ["customer_id", "customer-id"]));
    }

    // 識別子の同一性に関する残りのレビュー用メモ。
    // プロパティ同士の@・書式文字の比較は上の有効テストで確認済みです。
    // 以下は比較相手と非衝突の境界で分類し、項目種別×表記の直積にはしません。
    // 異種項目の集約は既存テストで確認済みで、Excel固有のSheet・Table・行データの区別はありません。

    // 1. 予約メンバーとの比較：予約名を全列挙せず、Readを代表にします。
    [Theory]
    [InlineData("@Read")]
    [InlineData("Re\u200Cad")]
    public void 表記が異なってもReadと同じ識別子になる生成プロパティ名は診断します(string propertyName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings["customerName"] = propertyName)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "Read", ["customerName"]));
    }

    // 2. 生成型名との比較：型名側にも書式文字がある場合を確認します。
    // DocumentとDataの両方へプロパティを生成するので、比較相手はこの二つです。
    [Theory]
    [InlineData("OrderDocument")]
    [InlineData("OrderData")]
    public void 書式文字を除くと生成型名と同じになるプロパティ名は診断します(string propertyName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings = new()
                {
                    ["customerData"] = "Or\u200Cder",
                    ["customerName"] = propertyName
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, propertyName, ["customerName"]));
    }

    [Fact]
    public void 生成Document型名にエスケープ表記があっても同じ識別子のプロパティを診断します()
    {
        // @OrderDocumentとOrderDocumentは、C#では同じ型名です。
        // プロパティ側ではなく、比較相手の生成型名に@を付けます。
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\customerData.docx",
                options => options.NameMappings = new()
                {
                    ["customerData"] = "@Order",
                    ["customerName"] = "OrderDocument"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "OrderDocument", ["customerName"]));
    }

    // 3. 非衝突の境界：文字の大小は区別し、利用者が衝突を解消できることを確認します。
    [Fact]
    public void 大文字小文字だけが異なる生成プロパティ名は衝突にはなりません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options => options.NameMappings = new()
                {
                    ["customer_id"] = "SharedValue",
                    ["customer-id"] = "sharedValue"
                })
            .Should().BeEmpty();
    }

    [Fact]
    public void 書式文字を含む生成名との衝突をNameMappingsで解消できます()
    {
        // 書式文字入りの名前は残し、もう一方だけを別の識別子に変更します。
        GeneratedCodeInspection
            .GenerateDiagnostics(
                MergeFieldNameCollisionDocumentFilePath,
                options => options.NameMappings = new()
                {
                    ["customer_id"] = "Shared\u200CValue",
                    ["customer-id"] = "OtherValue"
                })
            .Should().BeEmpty();
    }
}
