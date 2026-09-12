using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class MSBuild連携タスクのテスト
{
    [Fact]
    public void SDK形式プロジェクトのビルドアクション候補にDocumentAsDataを登録します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        Directory.CreateDirectory(project.DirectoryPath);
        var scriptFilePath = project.AddPowerShellSdkProjectSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "AvailableItemNames.txt"))
            .Should().Contain("DocumentAsData");
    }

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
    public void SDK形式プロジェクトの通常ビルドでコードを自動生成してコンパイルできます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureDocument("BasicStructure.docx");
        var scriptFilePath = project.AddPowerShellSdkBuildSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllText(project.GeneratedFilePathFor("BasicStructure.docx"))
            .Should().Contain("public partial class BasicStructureDocument : Document");
    }

    [Fact]
    public void 生成コードは不要なnewの警告をエラー扱いにしてもビルドできます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureDocument("BasicStructure.docx");
        var scriptFilePath = project.AddPowerShellSdkBuildSample(warningsAsErrors: "CS0109");

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
    }

    [Fact]
    public void SDK形式プロジェクトのRootNamespaceで生成型を利用できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureDocument("BasicStructure.docx");
        var scriptFilePath = project.AddPowerShellSdkBuildSample(rootNamespace: "Sample.Documents");

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllText(project.GeneratedFilePathFor("BasicStructure.docx"))
            .Should().Contain("namespace Sample.Documents;");
    }

    [Fact]
    public void 対象Word文書がないSDK形式プロジェクトも通常ビルドできます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var scriptFilePath = project.AddPowerShellSdkBuildSample(includeDocument: false);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
    }

    [Fact]
    public void デザイン時のコンパイルでは辞書が変更されてもコードを再生成しません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddCustomerDataDocument("BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var generatedSource = File.ReadAllText(project.GeneratedFilePathFor("BasicStructure.docx"));
        project.AddProjectDictionaryFor(
            "BasicStructure.docx",
            """
            {
              "customerName": "ClientName"
            }
            """);
        var scriptFilePath = project.AddPowerShellSdkBuildSample(designTimeBuild: true);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllText(project.GeneratedFilePathFor("BasicStructure.docx"))
            .Should().Be(generatedSource);
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
    public void 生成対象から外した文書の生成コードはコンパイル対象に含めません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var generatedFilePath = project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx");
        var scriptFilePath = project.AddPowerShellSdkProjectSample(includeDocument: false);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().NotContain(generatedFilePath);
        File.Exists(generatedFilePath).Should().BeTrue();
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
    public void SDK形式プロジェクトのDesignTimeBuildでは元Word文書に生成ファイル名を設定します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample();

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "DocumentGeneratedOutput.txt"))
            .Should().Equal($"{documentFilePath}|BasicStructure.DocumentAsData.g.cs");
    }

    [Fact]
    public void SDK形式プロジェクトのDesignTimeBuildでは生成コードをNone項目と重複させません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample(includeGeneratedSourceAsNone: true);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().Equal(project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"));
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "None.txt"))
            .Should().NotContain(project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"))
            .And.Contain(documentFilePath);
    }

    [Fact]
    public void SDK形式プロジェクトのDesignTimeBuildでは生成コードをContent項目と重複させません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample(includeGeneratedSourceAsContent: true);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().Equal(project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx"));
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Content.txt"))
            .Should().Equal(documentFilePath);
    }

    [Fact]
    public void SDK形式プロジェクトの通常評価でも生成コードを元Word文書へ紐づけます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var scriptFilePath = project.AddPowerShellSdkProjectSample(designTimeBuild: false);

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
    public void 生成内容が同じ場合は生成ファイルの更新日時を変更しません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var generatedFilePath = project.GeneratedFilePathFor(@"Schemas\BasicStructure.docx");
        var lastWriteTime = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        // 現在時刻との差を確実に作り、ファイルシステムの時刻精度を待機で補わないようにします。
        File.SetLastWriteTimeUtc(generatedFilePath, lastWriteTime);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        File.GetLastWriteTimeUtc(generatedFilePath).Should().Be(lastWriteTime);
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
    public void 辞書を変更して再生成すると生成プロパティ名を更新します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddCustomerDataDocument(@"Schemas\customerData.docx");
        project.AddProjectDictionaryFor(
            "customerData.docx",
            """
            {
              "customerName": "ClientName"
            }
            """);
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        File.ReadAllText(project.GeneratedFilePathFor(@"Schemas\customerData.docx"))
            .Should().Contain("public string ClientName");

        project.AddProjectDictionaryFor(
            "customerData.docx",
            """
            {
              "customerName": "RecipientName"
            }
            """);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource.Should().Contain("public string RecipientName")
            .And.NotContain("public string ClientName");
    }

    [Fact]
    public void 辞書を削除して再生成すると辞書なしのプロパティ名へ戻ります()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddCustomerDataDocument(@"Schemas\customerData.docx");
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        var sourceWithoutDictionary = File.ReadAllText(
            project.GeneratedFilePathFor(@"Schemas\customerData.docx"));
        var dictionaryFilePath = project.AddProjectDictionaryFor(
            "customerData.docx",
            """
            {
              "customerName": "ClientName"
            }
            """);
        project.Generate(documentFilePath).Succeeded.Should().BeTrue();
        File.ReadAllText(project.GeneratedFilePathFor(@"Schemas\customerData.docx"))
            .Should().Contain("public string ClientName");

        File.Delete(dictionaryFilePath);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource.Should().Be(sourceWithoutDictionary);
    }

    [Fact]
    public void 生成プロパティ名が衝突した場合は対象文書を示して失敗します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddMergeFieldNameCollisionDocument(@"Schemas\NameCollision.docx");

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeFalse();
        tested.Errors.Should().ContainSingle();
        tested.Errors.Single().File.Should().Be(documentFilePath);
        tested.Errors.Single().Message.Should().Contain("CustomerId")
            .And.Contain("customer_id").And.Contain("customer-id");
        File.Exists(project.GeneratedFilePathFor(@"Schemas\NameCollision.docx"))
            .Should().BeFalse();
    }

    [Fact]
    public void 辞書で生成プロパティ名の衝突を解消するとコード生成できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddMergeFieldNameCollisionDocument(@"Schemas\NameCollision.docx");
        project.AddDocumentDictionaryFor(
            @"Schemas\NameCollision.docx",
            """
            {
              "customer-id": "CustomerIdDash"
            }
            """);

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.Errors.Should().BeEmpty();
        tested.SingleGeneratedSource.Should().Contain("public string CustomerIdDash");
    }

    [Fact]
    public void 存在しないWord文書を指定すると対象文書を示して失敗します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        Directory.CreateDirectory(project.DirectoryPath);
        var documentFilePath = Path.Combine(project.DirectoryPath, "Missing.docx");

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeFalse();
        tested.Errors.Should().ContainSingle();
        tested.Errors.Single().File.Should().Be(documentFilePath);
        tested.Errors.Single().Message.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void 不正な辞書JSONでは辞書ファイルを示して失敗します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var documentFilePath = project.AddBasicStructureDocument(@"Schemas\BasicStructure.docx");
        var dictionaryFilePath = project.AddDocumentDictionaryFor(
            @"Schemas\BasicStructure.docx",
            "{");

        var tested = project.Generate(documentFilePath);

        tested.Succeeded.Should().BeFalse();
        tested.Errors.Should().ContainSingle();
        tested.Errors.Single().File.Should().Be(dictionaryFilePath);
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
