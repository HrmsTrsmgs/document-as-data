using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成統合のテスト
{
    const string IntegratedDocumentFilePath = @"TestData\コード生成\統合.docx";

    [Fact]
    public void 指定した名前空間へすべての型を生成します()
    {
        var generatedTypes = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    IntegratedDocumentFilePath,
                    options => options.Namespace = "Generated.Custom"))
            .DefinedTypes
            .ToArray();

        generatedTypes
            .Should().NotBeEmpty();

        generatedTypes
            .Should().OnlyContain(it => it.Namespace == "Generated.Custom");
    }
}
