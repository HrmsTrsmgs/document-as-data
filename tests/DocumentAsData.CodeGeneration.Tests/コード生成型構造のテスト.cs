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
