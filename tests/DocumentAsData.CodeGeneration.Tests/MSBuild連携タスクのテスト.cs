using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class MSBuild連携タスクのテスト
{
    [Fact]
    public void DocumentAsData項目からWord文書の隣へ生成コードを出力します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath =
            project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.Warnings.Should().BeEmpty();
        tested.GeneratedFilePaths.Should().ContainSingle();
        tested.SingleGeneratedFilePath
            .Should().Be(
                project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"));
        tested.SingleGeneratedSource
            .Should().Contain(
                "public partial class BasicStructureDocument : Document");
    }
}
