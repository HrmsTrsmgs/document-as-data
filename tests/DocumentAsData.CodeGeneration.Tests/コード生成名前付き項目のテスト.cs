using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前付き項目のテスト : IDisposable
{
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";
    const string BackslashMergeFieldNameDocumentFilePath =
        @"TestData\コード生成\バックスラッシュを含むMERGEFIELD名.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string SpecialCharacterTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTag.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";
    const string SpecialCharacterCheckBoxTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTagのCheckBox.docx";
    const string DatePickerDocumentFilePath =
        @"TestData\コード生成\日付選択ContentControl.docx";
    const string SpecialCharacterDatePickerTagDocumentFilePath =
        @"TestData\コード生成\引用符とバックスラッシュを含むTagのDatePicker.docx";

    readonly TemporaryDocumentFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldsDocumentFilePath))
            .GeneratedType("MergefieldDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                MergeFieldsDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void バックスラッシュを含む名前から生成したMERGEFIELDプロパティで文字列を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BackslashMergeFieldNameDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp customer"] = "SpecialValue"))
            .GeneratedType("バックスラッシュを含むMERGEFIELD名Document")
            .InvokeStaticMethod<Document>(
                "Open",
                BackslashMergeFieldNameDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.SpecialValue;

        tested.Should().Be("株式会社○○");
    }

    [Fact]
    public void 生成されたDocument型のMERGEFIELDプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           MergeFieldsDocumentFilePath))
                   .GeneratedType("MergefieldDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       MergeFieldsDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.MergeFields["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティから文字列を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                TextContentControlDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.CustomerName;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 特殊文字を含むTagから生成したプロパティで文字列を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""document"] = "SpecialValue"))
            .GeneratedType("引用符とバックスラッシュを含むTagDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterTagDocumentFilePath);

        dynamic documentAccessor = document;
        string tested = documentAccessor.SpecialValue;

        tested.Should().Be("山田太郎");
    }

    [Fact]
    public void 生成されたDocument型の文字列ContentControlプロパティへ文字列を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           TextContentControlDocumentFilePath))
                   .GeneratedType("文字列ContentControlDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       TextContentControlDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.CustomerName = "生成後";
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.ContentControls["CustomerName"].Text.Should().Be("生成後");
    }

    [Fact]
    public void 生成されたDocument型のCheckBoxプロパティからチェック状態を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.Agreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 生成されたDocument型のReadでCheckBoxを文書データとして読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                CheckBoxDocumentFilePath);

        dynamic documentAccessor = document;
        var data = documentAccessor.Read();
        bool tested = data.Agreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 特殊文字を含むTagから生成したCheckBoxプロパティでチェック状態を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterCheckBoxTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""agreement"] = "SpecialAgreement"))
            .GeneratedType("引用符とバックスラッシュを含むTagのCheckBoxDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterCheckBoxTagDocumentFilePath);

        dynamic documentAccessor = document;
        bool tested = documentAccessor.SpecialAgreement;

        tested.Should().BeTrue();
    }

    [Fact]
    public void 生成されたDocument型のCheckBoxプロパティへチェック状態を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           CheckBoxDocumentFilePath))
                   .GeneratedType("チェック済みCheckBoxDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       CheckBoxDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.Agreement = false;
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);

        tested.CheckBoxes["Agreement"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void 生成されたDocument型のDatePickerプロパティから日時を直接読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath))
            .GeneratedType("日付選択ContentControlDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                DatePickerDocumentFilePath);

        dynamic documentAccessor = document;
        DateTimeOffset tested = documentAccessor.DeliveryDate;

        tested.Should().Be(
            new DateTimeOffset(
                2026,
                9,
                4,
                0,
                0,
                0,
                TimeSpan.Zero));
    }

    [Fact]
    public void 特殊文字を含むTagから生成したDatePickerプロパティで日時を読み取れます()
    {
        using var document = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SpecialCharacterDatePickerTagDocumentFilePath,
                    options =>
                        options.NameMappings[
                            @"C:\temp\""delivery"] = "SpecialDeliveryDate"))
            .GeneratedType("引用符とバックスラッシュを含むTagのDatePickerDocument")
            .InvokeStaticMethod<Document>(
                "Open",
                SpecialCharacterDatePickerTagDocumentFilePath);

        dynamic documentAccessor = document;
        DateTimeOffset tested = documentAccessor.SpecialDeliveryDate;

        tested.Should().Be(
            new DateTimeOffset(
                2026,
                9,
                4,
                0,
                0,
                0,
                TimeSpan.Zero));
    }

    [Fact]
    public void 生成されたDocument型のDatePickerプロパティへ日時を直接書き込めます()
    {
        var savedPath = temporaryFiles.NewFilePath();
        var value = new DateTimeOffset(
            2026,
            12,
            31,
            0,
            0,
            0,
            TimeSpan.FromHours(9));

        using (var document = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DatePickerDocumentFilePath))
                   .GeneratedType("日付選択ContentControlDocument")
                   .InvokeStaticMethod<Document>(
                       "Open",
                       DatePickerDocumentFilePath))
        {
            dynamic documentAccessor = document;

            documentAccessor.DeliveryDate = value;
            document.SaveAs(savedPath);
        }

        using var tested = Document.Open(savedPath);
        var actual = tested.DatePickers["DeliveryDate"].SelectedDateTime;

        actual.Should().Be(value);
        actual.Offset.Should().Be(value.Offset);
    }
}
