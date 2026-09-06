using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成型構造のテスト
{
    const string BasicStructureDocumentFilePath = @"TestData\コード生成\BasicStructure.docx";
    const string IntegratedDocumentFilePath = @"TestData\コード生成\統合.docx";

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
}
