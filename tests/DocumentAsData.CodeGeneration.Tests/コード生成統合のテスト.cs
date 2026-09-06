using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成統合のテスト
{
    const string IntegratedDocumentFilePath = @"TestData\コード生成\統合.docx";

    [Fact]
    public void 基本的なWord文書からコンパイルできるソースを生成します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    IntegratedDocumentFilePath))
            .Should().NotBeNull();
    }

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

    [Fact]
    public void 同じWord文書と設定から同じ生成結果を返します()
    {
        GeneratedCodeInspection
            .GenerateSources(IntegratedDocumentFilePath)
            .Should().Equal(GeneratedCodeInspection.GenerateSources(
                IntegratedDocumentFilePath));
    }

    [Fact]
    public void 正常なWord文書ではエラー診断を返しません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(IntegratedDocumentFilePath)
            .Should().BeEmpty();
    }
}
