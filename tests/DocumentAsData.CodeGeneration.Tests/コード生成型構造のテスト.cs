using System.Reflection;
using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成型構造のテスト
{
    const string BasicStructureDocumentFilePath = @"TestData\コード生成\BasicStructure.docx";
    const string IntegratedDocumentFilePath = @"TestData\コード生成\統合.docx";
    const string MergeFieldDocumentFilePath = @"TestData\コード生成\MERGEFIELD.docx";
    const string TextContentControlDocumentFilePath =
        @"TestData\コード生成\文字列ContentControl.docx";
    const string CheckBoxDocumentFilePath =
        @"TestData\コード生成\チェック済みCheckBox.docx";
    const string DatePickerDocumentFilePath =
        @"TestData\コード生成\日付選択ContentControl.docx";

    [Theory]
    [InlineData(BasicStructureDocumentFilePath, "BasicStructureDocument")]
    [InlineData(IntegratedDocumentFilePath, "統合Document")]
    public void 生成されたDocument型はWord文書ファイル名に対応する型名で生成します(
        string documentFilePath,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(documentFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain(generatedTypeName);
    }

    [Fact]
    public void 生成されたDocument型はDocumentを継承します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureDocumentFilePath))
            .GeneratedType("BasicStructureDocument")
            .Should().BeAssignableTo<Document>();
    }

    [Fact]
    public void Word文書全体のデータを表す型を生成します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureDocumentFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain("BasicStructureData");
    }

    [Theory]
    [InlineData(
        MergeFieldDocumentFilePath,
        "MergefieldData")]
    [InlineData(
        TextContentControlDocumentFilePath,
        "文字列ContentControlData")]
    public void 文字列項目を文書データ型のstringプロパティとして生成します(
        string documentFilePath,
        string dataTypeName)
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(documentFilePath))
            .GeneratedType(dataTypeName)
            .GetProperty("CustomerName");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(string));
    }

    [Fact]
    public void CheckBoxを文書データ型のboolプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    CheckBoxDocumentFilePath))
            .GeneratedType("チェック済みCheckBoxData")
            .GetProperty("Agreement");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(bool));
    }

    [Fact]
    public void DatePickerを文書データ型のDateTimeOffsetプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DatePickerDocumentFilePath))
            .GeneratedType("日付選択ContentControlData")
            .GetProperty("DeliveryDate");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(DateTimeOffset));
    }

    [Fact]
    public void 生成されたDocument型は文書データ型を返すReadメソッドを公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(
                BasicStructureDocumentFilePath));
        var dataType = generatedAssembly.GeneratedType("BasicStructureData");
        var tested = generatedAssembly
            .GeneratedType("BasicStructureDocument")
            .GetMethods(
                BindingFlags.Public |
                BindingFlags.Instance |
                BindingFlags.DeclaredOnly)
            .SingleOrDefault(it =>
                it.Name == "Read" &&
                it.GetParameters().Length == 0);

        tested.Should().NotBeNull();
        tested.IsGenericMethod.Should().BeFalse();
        tested.ReturnType.Should().Be(dataType);
    }

    [Fact]
    public void 生成されたDocument型は文書データ型を受け取るReplaceメソッドを公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(
                BasicStructureDocumentFilePath));
        var dataType = generatedAssembly.GeneratedType("BasicStructureData");
        var tested = generatedAssembly
            .GeneratedType("BasicStructureDocument")
            .GetMethods(
                BindingFlags.Public |
                BindingFlags.Instance |
                BindingFlags.DeclaredOnly)
            .SingleOrDefault(it =>
                it.Name == "Replace" &&
                it.GetParameters().Select(parameter => parameter.ParameterType)
                    .SequenceEqual([dataType]));

        tested.Should().NotBeNull();
        tested.ReturnType.Should().Be(typeof(void));
    }

    [Fact]
    public void 生成されたDocument型は指定ファイルを開く静的Openメソッドを公開します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureDocumentFilePath))
            .GeneratedType("BasicStructureDocument")
            .GetMethod("Open", [typeof(string)]);

        tested.Should().NotBeNull();
        tested.IsStatic.Should().BeTrue();
        tested.ReturnType.Should().Be(
            tested.DeclaringType);
    }

    [Fact]
    public void 生成されたDocument型は必要なMERGEFIELDが不足した文書をOpenすると例外になります()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldDocumentFilePath))
            .GeneratedType("MergefieldDocument");

        // 生成元にはCustomerNameとAddressがあり、開く文書には名前付き項目がありません。
        var tested = () =>
        {
            using var document = generatedType.InvokeStaticMethod<Document>(
                "Open",
                BasicStructureDocumentFilePath);
        };

        // リフレクションによる呼び出しでは、Openの例外がInnerExceptionに入ります。
        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<DocumentMappingException>();
    }

    [Fact]
    public void 生成されたDocument型は必要な文字列ContentControlが不足した文書をOpenすると例外になります()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    TextContentControlDocumentFilePath))
            .GeneratedType("文字列ContentControlDocument");

        // 生成元にはTagがCustomerNameの文字列Content Controlがあり、開く文書にはありません。
        var tested = () =>
        {
            using var document = generatedType.InvokeStaticMethod<Document>(
                "Open",
                BasicStructureDocumentFilePath);
        };

        // リフレクションによる呼び出しでは、Openの例外がInnerExceptionに入ります。
        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<DocumentMappingException>();
    }

    [Fact]
    public void 生成されたDocument型はStreamから開けます()
    {
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(BasicStructureDocumentFilePath));
        stream.Position = 0;
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureDocumentFilePath))
            .GeneratedType("BasicStructureDocument");

        using var tested =
            generatedType.InvokeStaticMethod<Document>("Open", stream);

        tested.GetType().Should().Be(generatedType);
    }

    [Fact]
    public void 生成されたDocument型は必要なMERGEFIELDが不足したStreamをOpenすると例外になります()
    {
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(BasicStructureDocumentFilePath));
        stream.Position = 0;
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldDocumentFilePath))
            .GeneratedType("MergefieldDocument");

        // 生成元にはCustomerNameとAddressがあり、Stream上の文書には名前付き項目がありません。
        var tested = () =>
        {
            using var document = generatedType.InvokeStaticMethod<Document>(
                "Open",
                stream);
        };

        // リフレクションによる呼び出しでは、Openの例外がInnerExceptionに入ります。
        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<DocumentMappingException>();
    }

    [Fact]
    public void 生成されたDocument型は別ファイルのpartial定義と共にコンパイルできます()
    {
        GeneratedSourceCompiler
            .Compile(
                [
                    .. GeneratedCodeInspection.GenerateSources(
                        BasicStructureDocumentFilePath),
                    """
                    namespace Generated;

                    public partial class BasicStructureDocument
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GeneratedType("BasicStructureDocument")
            .GetProperty("AddedByUser")
            .Should().NotBeNull();
    }
}
