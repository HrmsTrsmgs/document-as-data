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

    [Fact]
    public void 生成コードは元Word文書へ紐づくメタデータを返します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath =
            project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedFile.GetMetadata("DependentUpon")
            .Should().Be("BasicStructure.docx");
        tested.SingleGeneratedFile.GetMetadata("DesignTimeSharedInput")
            .Should().Be("true");
    }

    [Fact]
    public void プロジェクト直下の辞書を使用します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath =
            project.AddCustomerDataDocument(@"Schemas\customerData.docx");
        project.AddProjectDictionaryFor(
            "customerData.docx",
            """
            {
              "customerName": "ClientName"
            }
            """);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public string ClientName");
    }
}
