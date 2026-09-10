using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class MSBuild連携タスクのテスト
{
    [Fact]
    public void PowerShellからdotnet_msbuildでコード生成できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureDocument("BasicStructure.docx");
        var scriptFilePath = project.AddPowerShellGenerationSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllText(project.GeneratedFilePathFor("BasicStructure.docx"))
            .Should().Contain("public partial class BasicStructureDocument : Document");
    }

    [Fact]
    public void PowerShellからdotnet_msbuildのDesignTimeBuildで生成コードを参照できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellDesignTimeBuildSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().Equal(project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"));
    }

    [Fact]
    public void SDK形式プロジェクトのDesignTimeBuildでは生成コードを既定Compileと重複させません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().Equal(project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"));
    }

    [Fact]
    public void SDK形式プロジェクトのDesignTimeBuildでは生成コードを元Word文書へ紐づけます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "CompileNesting.txt"))
            .Should().Equal(
                $"{project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx")}|BasicStructure.docx");
    }

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

    [Fact]
    public void Word文書と同じディレクトリの辞書をプロジェクト直下の辞書より優先します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath =
            project.AddCustomerDataDocument(@"Schemas\customerData.docx");
        project.AddProjectDictionaryFor(
            "customerData.docx",
            """
            {
              "customerName": "ProjectRootClientName"
            }
            """);
        project.AddDocumentDictionaryFor(
            @"Schemas\customerData.docx",
            """
            {
              "customerName": "SameDirectoryClientName"
            }
            """);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public string SameDirectoryClientName");
    }

    [Fact]
    public void 異なるディレクトリにある同名Word文書の生成ファイルは衝突しません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var orderFilePath =
            project.AddBasicStructureDocument(@"Orders\Master.docx");
        var archiveFilePath =
            project.AddBasicStructureDocument(@"Archive\Master.docx");

        var tested = project.Generate(orderFilePath, archiveFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.GeneratedFilePaths
            .Should().BeEquivalentTo(
                [
                    project.GeneratedFilePathFor(@"Orders\Master.docx"),
                    project.GeneratedFilePathFor(@"Archive\Master.docx")
                ]);
    }
}
