using System.Collections;
using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Marimo.DocumentAsData.Build;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// MSBuild連携のテストで、利用者プロジェクトに近い一時ディレクトリを扱います。
/// </summary>
sealed class MSBuild連携テストプロジェクト : IDisposable
{
    const string BasicStructureDocumentFilePath =
        @"TestData\コード生成\BasicStructure.docx";
    const string CustomerDataDocumentFilePath =
        @"TestData\コード生成\customerData.docx";

    /// <summary>
    /// テストごとに分離した一時プロジェクトを作成します。
    /// </summary>
    MSBuild連携テストプロジェクト()
    {
        DirectoryPath = Path.Combine(
            Path.GetTempPath(),
            "DocumentAsData.Tests",
            Guid.NewGuid().ToString("N"));
    }

    /// <summary>
    /// テスト対象プロジェクトのルートディレクトリです。
    /// </summary>
    internal string DirectoryPath { get; }

    /// <summary>
    /// パッケージ配置に依存せず、リポジトリで開発中のtargetsを参照します。
    /// </summary>
    static string TargetsPath => Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory,
        "..", "..", "..", "..", "..",
        "src", "DocumentAsData.Build", "buildTransitive",
        "Marimo.DocumentAsData.Build.targets"));

    /// <summary>
    /// テストと同じ構成でビルド済みのタスクを、一時ディレクトリへパッケージ化します。
    /// restoreと再ビルドは行わず、パッケージの公開もしません。
    /// </summary>
    /// <returns>パッケージを作成するPowerShellスクリプトのパス。</returns>
    /// <param name="includeProjectReferences">利用側のrestoreに必要な参照プロジェクトも梱包する場合はtrue。</param>
    /// <param name="warningsAsErrors">梱包時の警告も失敗として検証する場合はtrue。</param>
    /// <param name="includeCombinedPackage">統合パッケージも梱包する場合はtrue。</param>
    internal string AddPowerShellPackSample(bool includeProjectReferences = false, bool warningsAsErrors = false, bool includeCombinedPackage = false)
    {
        Directory.CreateDirectory(DirectoryPath);
        var projectFilePath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "src", "DocumentAsData.Build", "DocumentAsData.Build.csproj"));
        var configuration = typeof(GenerateDocumentAsData).Assembly
            .GetCustomAttributes<AssemblyConfigurationAttribute>().Single().Configuration;
        string[] projectNames = includeProjectReferences
            ? ["DocumentAsData", "DocumentAsData.CodeGeneration", "DocumentAsData.Build"]
            : ["DocumentAsData.Build"];
        if (includeCombinedPackage)
            projectNames = [.. projectNames, "DocumentAsData.Package"];
        return WritePowerShellScript(
            "Pack.ps1",
            $$"""
            foreach ($projectName in @({{string.Join(", ", projectNames.Select(it => $"'{it}'"))}})) {
                $projectPath = Join-Path '{{Path.GetFullPath(Path.Combine(projectFilePath, "..", "..")).Replace("'", "''")}}' "$projectName/$projectName.csproj"
                dotnet pack $projectPath --no-build --no-restore --disable-build-servers --configuration '{{configuration.Replace("'", "''")}}' --output ./packages --nologo --verbosity minimal -p:TreatWarningsAsErrors={{warningsAsErrors}}
                if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
            }
            exit 0
            """);
    }

    /// <summary>
    /// 作成済みパッケージを参照し、明示的なImportなしで項目評価または通常ビルドを実行します。
    /// 復元先をテスト内へ分離し、外部依存は本リポジトリの復元済みキャッシュから取得します。
    /// DOCUMENTASDATA_TEST_MSBUILDを指定すると、そのMSBuild.exeで同じ利用者テストを実行できます。
    /// </summary>
    /// <param name="build">生成型を使うコードを通常ビルドする場合はtrue、項目評価だけならfalse。</param>
    /// <param name="buildTarget">buildがtrueの場合に実行するターゲット。Cleanも同じ利用者プロジェクトで検証します。</param>
    /// <param name="run">生成型で読み書きする実行用サンプルも作り、ビルド後に実行する場合はtrue。</param>
    /// <param name="designTimeBuild">文書を再生成しないデザイン時のコンパイルを確認する場合はtrue。</param>
    /// <param name="includeDocument">Word文書をコード生成対象として登録する場合はtrue。</param>
    /// <param name="packageFilePattern">利用者が参照するパッケージを選ぶファイル名パターン。</param>
    /// <param name="sharedLibrary">生成対象を共用ライブラリに置き、別のアプリからProjectReferenceする場合はtrue。</param>
    /// <returns>restoreと指定した検証を実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellPackageReferenceSample(
        bool build = false,
        string buildTarget = "Build",
        bool run = false,
        bool designTimeBuild = false,
        bool includeDocument = true,
        string packageFilePattern = "*DocumentAsData.Build.*.nupkg",
        bool sharedLibrary = false)
    {
        using var package = ZipFile.OpenRead(
            Directory.GetFiles(Path.Combine(DirectoryPath, "packages"), packageFilePattern).Single());
        using var manifestStream = package.Entries.Single(it => it.FullName.EndsWith(".nuspec")).Open();
        var manifest = XDocument.Load(manifestStream);
        var packageId = manifest.Descendants().Single(it => it.Name.LocalName == "id").Value;
        var packageVersion = manifest.Descendants().Single(it => it.Name.LocalName == "version").Value;
        using var assets = JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "..",
            "src", "DocumentAsData.Build", "obj", "project.assets.json"))));
        File.WriteAllText(
            Path.Combine(DirectoryPath, "NuGet.Config"),
            new XElement("configuration",
                new XElement("packageSources",
                    new XElement("clear"),
                    new XElement("add", new XAttribute("key", "local"), new XAttribute("value", "packages")),
                    assets.RootElement.GetProperty("packageFolders").EnumerateObject().Select((it, index) =>
                        new XElement("add", new XAttribute("key", $"cache{index}"), new XAttribute("value", it.Name)))))
                .ToString());
        File.WriteAllText(
            Path.Combine(DirectoryPath, "PackageReference.csproj"),
            $$"""
            <Project Sdk="Microsoft.NET.Sdk">
              <PropertyGroup>
                <TargetFramework>net10.0</TargetFramework>
                <OutputType Condition="'{{run}}' == 'True' and '{{sharedLibrary}}' != 'True'">Exe</OutputType>
                <DesignTimeBuild Condition="'{{designTimeBuild}}' == 'True'">true</DesignTimeBuild>
                <SkipCompilerExecution Condition="'{{designTimeBuild}}' == 'True'">true</SkipCompilerExecution>
                <RestorePackagesPath>$(MSBuildProjectDirectory)/restored</RestorePackagesPath>
                <NuGetAudit>false</NuGetAudit>
                <RootNamespace>Generated</RootNamespace>
                <UseSharedCompilation>false</UseSharedCompilation>
              </PropertyGroup>
              <ItemGroup>
                <PackageReference Include="{{SecurityElement.Escape(packageId)}}" Version="{{SecurityElement.Escape(packageVersion)}}" />
                <DocumentAsData Include="BasicStructure.docx" Condition="'{{build}}' == 'True' and '{{includeDocument}}' == 'True'" />
                <Compile Remove="App/**/*.cs" />
              </ItemGroup>
              <Target Name="WriteAvailableItems">
                <WriteLinesToFile File="AvailableItemNames.txt" Lines="@(AvailableItemName)" Overwrite="true" />
              </Target>
              <Target Name="WriteProjectItems">
                <!-- DependentUponは生成コードの親文書、LastGenOutputは文書側から見た生成ファイル名です。 -->
                <WriteLinesToFile File="CompileNesting.txt" Lines="@(Compile->'%(FullPath)|%(DependentUpon)')" Overwrite="true" />
                <WriteLinesToFile File="DocumentGeneratedOutput.txt" Lines="@(DocumentAsData->'%(FullPath)|%(LastGenOutput)')" Overwrite="true" />
                <WriteLinesToFile File="OtherItems.txt" Lines="@(None);@(Content)" Overwrite="true" />
              </Target>
              <!-- ReferencePathは、コンパイラへ渡すために解決済みのアセンブリ参照です。 -->
              <Target Name="WriteResolvedReferences" AfterTargets="ResolveReferences">
                <WriteLinesToFile File="References.txt" Lines="@(ReferencePath->'%(Filename)%(Extension)')" Overwrite="true" />
              </Target>
            </Project>
            """);
        if (sharedLibrary)
        {
            Directory.CreateDirectory(Path.Combine(DirectoryPath, "App"));
            File.WriteAllText(
                Path.Combine(DirectoryPath, "App", "App.csproj"),
                """
                <Project Sdk="Microsoft.NET.Sdk">
                  <PropertyGroup>
                    <TargetFramework>net10.0</TargetFramework>
                    <OutputType>Exe</OutputType>
                    <RestorePackagesPath>$(MSBuildProjectDirectory)/../restored</RestorePackagesPath>
                    <NuGetAudit>false</NuGetAudit>
                    <UseSharedCompilation>false</UseSharedCompilation>
                  </PropertyGroup>
                  <ItemGroup>
                    <ProjectReference Include="../PackageReference.csproj" />
                  </ItemGroup>
                </Project>
                """);
        }
        if (build)
        {
            File.WriteAllText(
                Path.Combine(DirectoryPath, "Consumer.cs"),
                """
                public sealed class Consumer
                {
                    public Generated.BasicStructureDocument Document { get; set; }
                }
                """);
        }
        if (run)
        {
            File.WriteAllText(
                Path.Combine(DirectoryPath, sharedLibrary ? "App" : "", "Program.cs"),
                """
                using System;
                using Generated;

                using (var document = BasicStructureDocument.Open("BasicStructure.docx"))
                {
                    var data = document.Read();
                    Console.WriteLine($"before:{data.CustomerName}/{data.Address}");
                    data.CustomerName = "更新後";
                    data.Address = "大阪府";
                    document.Replace(data);
                    document.SaveAs("output.docx");
                }
                using var saved = BasicStructureDocument.Open("output.docx");
                Console.WriteLine($"after:{saved.CustomerName}/{saved.Address}");
                """);
        }
        var consumerProject = sharedLibrary ? "App/App.csproj" : "PackageReference.csproj";
        return WritePowerShellScript(
            "PackageReference.ps1",
            $$"""
            dotnet restore ./{{consumerProject}} --configfile ./NuGet.Config --nologo --verbosity minimal
            if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
            $buildArguments = @('./{{consumerProject}}', '/t:{{(build
                ? buildTarget
                : "WriteAvailableItems")}}', '/nologo', '/v:minimal', '/nr:false')
            if ($env:DOCUMENTASDATA_TEST_MSBUILD) {
                & $env:DOCUMENTASDATA_TEST_MSBUILD @buildArguments
            } else {
                dotnet msbuild @buildArguments
            }
            if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
            if ('{{run}}' -eq 'True') {
                dotnet run --project ./{{consumerProject}} --no-build --no-restore
            }
            exit $LASTEXITCODE
            """);
    }

    /// <summary>
    /// リポジトリのtargetsとビルド済みタスクを使う一時プロジェクトを作ります。パッケージの配置は対象外です。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellGenerationSample()
    {
        File.WriteAllText(
            Path.Combine(DirectoryPath, "DocumentAsData.Generate.proj"),
            $$"""
            <Project>
              <PropertyGroup>
                <DocumentAsDataTaskAssembly>{{SecurityElement.Escape(typeof(GenerateDocumentAsData).Assembly.Location)}}</DocumentAsDataTaskAssembly>
              </PropertyGroup>
              <ItemGroup>
                <DocumentAsData Include="BasicStructure.docx" />
              </ItemGroup>
              <Import Project="{{SecurityElement.Escape(TargetsPath)}}" />
              <Target Name="Build" DependsOnTargets="GenerateDocumentAsDataSources" />
            </Project>
            """);
        return WritePowerShellScript(
            "Generate.ps1",
            """
            dotnet msbuild ./DocumentAsData.Generate.proj /t:Build /nologo /v:minimal
            exit $LASTEXITCODE
            """);
    }

    /// <summary>
    /// 生成型を参照するSDKプロジェクトを作り、コンパイル経路でのコード生成を観測します。
    /// ビルド済みライブラリを参照し、外部パッケージソースを使わずにrestoreします。
    /// </summary>
    /// <param name="designTimeBuild">デザイン時の生成抑止を確認する場合はtrue。</param>
    /// <param name="includeDocument">生成対象と生成型の参照を含める場合はtrue。</param>
    /// <param name="rootNamespace">プロジェクトに指定し、利用コードから生成型を参照する名前空間。</param>
    /// <param name="warningsAsErrors">生成コードの検証でエラーとして扱うコンパイラ警告コード。</param>
    /// <returns>ビルドを実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellSdkBuildSample(
        bool designTimeBuild = false,
        bool includeDocument = true,
        string rootNamespace = "Generated",
        string warningsAsErrors = "")
    {
        Directory.CreateDirectory(DirectoryPath);
        File.WriteAllText(
            Path.Combine(DirectoryPath, "DocumentAsData.BuildSample.csproj"),
            $$"""
            <Project Sdk="Microsoft.NET.Sdk">
              <PropertyGroup>
                <TargetFramework>net10.0</TargetFramework>
                <Nullable>enable</Nullable>
                <RootNamespace>{{SecurityElement.Escape(rootNamespace)}}</RootNamespace>
                <WarningsAsErrors>{{SecurityElement.Escape(warningsAsErrors)}}</WarningsAsErrors>
                <DesignTimeBuild>{{designTimeBuild}}</DesignTimeBuild>
                <UseSharedCompilation>false</UseSharedCompilation>
                <DocumentAsDataTaskAssembly>{{SecurityElement.Escape(typeof(GenerateDocumentAsData).Assembly.Location)}}</DocumentAsDataTaskAssembly>
              </PropertyGroup>
              <ItemGroup>
                <Reference Include="DocumentAsData" HintPath="{{SecurityElement.Escape(typeof(Document).Assembly.Location)}}" />
                <DocumentAsData Include="BasicStructure.docx" Condition="'{{includeDocument}}' == 'True'" />
              </ItemGroup>
              <Import Project="{{SecurityElement.Escape(TargetsPath)}}" />
            </Project>
            """);
        File.WriteAllText(
            Path.Combine(DirectoryPath, "Consumer.cs"),
            $$"""
            public sealed class Consumer
            {
                public {{(includeDocument
                    ? $"{rootNamespace}.BasicStructureDocument"
                    : "Marimo.DocumentAsData.Document")}}? Document { get; set; }
            }
            """);
        File.WriteAllText(
            Path.Combine(DirectoryPath, "NuGet.Config"),
            """
            <configuration>
              <packageSources><clear /></packageSources>
            </configuration>
            """);
        return WritePowerShellScript(
            "Build.ps1",
            """
            dotnet build ./DocumentAsData.BuildSample.csproj --nologo --verbosity minimal
            exit $LASTEXITCODE
            """);
    }

    /// <summary>
    /// 生成を実行せず、DesignTimeBuildで評価されたCompile項目を書き出す一時プロジェクトを作ります。
    /// SDKの既定Compile項目に助けられないよう、SDKを指定しないプロジェクトで登録を観測します。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellDesignTimeBuildSample()
    {
        File.WriteAllText(
            Path.Combine(DirectoryPath, "DocumentAsData.DesignTimeBuild.proj"),
            $$"""
            <Project>
              <PropertyGroup>
                <DesignTimeBuild>true</DesignTimeBuild>
                <DocumentAsDataTaskAssembly>{{SecurityElement.Escape(typeof(GenerateDocumentAsData).Assembly.Location)}}</DocumentAsDataTaskAssembly>
              </PropertyGroup>
              <ItemGroup>
                <DocumentAsData Include="Schemas\BasicStructure.docx" />
              </ItemGroup>
              <Import Project="{{SecurityElement.Escape(TargetsPath)}}" />
              <Target Name="Build">
                <WriteLinesToFile File="Compile.txt" Lines="@(Compile->'%(FullPath)')" Overwrite="true" />
              </Target>
            </Project>
            """);
        return WritePowerShellScript(
            "DesignTimeBuild.ps1",
            """
            dotnet msbuild ./DocumentAsData.DesignTimeBuild.proj /t:Build /nologo /v:minimal
            exit $LASTEXITCODE
            """);
    }

    /// <summary>
    /// SDKの既定Compile項目が有効なプロジェクトを作り、項目の登録結果を観測します。
    /// コンパイルやrestoreは行わず、生成済みソースの項目評価だけを確認します。
    /// </summary>
    /// <param name="includeGeneratedSourceAsNone">生成ソースをNoneにも明示登録し、重複除去を検証する場合はtrue。</param>
    /// <param name="includeGeneratedSourceAsContent">生成ソースと元文書をContentにも登録し、生成ソースだけの重複除去を検証する場合はtrue。</param>
    /// <param name="designTimeBuild">デザイン時の評価はtrue、通常の評価はfalse。</param>
    /// <param name="includeDocument">文書をコード生成対象として登録する場合はtrue。</param>
    /// <returns>実行するPowerShellスクリプトのパス。</returns>
    internal string AddPowerShellSdkProjectSample(
        bool includeGeneratedSourceAsNone = false,
        bool includeGeneratedSourceAsContent = false,
        bool designTimeBuild = true,
        bool includeDocument = true)
    {
        File.WriteAllText(
            Path.Combine(DirectoryPath, "DocumentAsData.SdkProject.csproj"),
            $$"""
            <Project Sdk="Microsoft.NET.Sdk">
              <Import Project="{{SecurityElement.Escape(Path.ChangeExtension(TargetsPath, ".props"))}}" />
              <PropertyGroup>
                <TargetFramework>net10.0</TargetFramework>
                <DesignTimeBuild>{{designTimeBuild}}</DesignTimeBuild>
                <DocumentAsDataTaskAssembly>{{SecurityElement.Escape(typeof(GenerateDocumentAsData).Assembly.Location)}}</DocumentAsDataTaskAssembly>
              </PropertyGroup>
              <ItemGroup>
                <!-- SDKは通常.csをNoneから除外するため、重複ケースでは明示的に追加します。 -->
                <None Include="Schemas\BasicStructure.DocumentAsData.g.cs" Condition="'{{includeGeneratedSourceAsNone}}' == 'True'" />
                <!-- 元文書もContentに登録し、生成ソース以外を除去しないことを確認します。 -->
                <Content Include="Schemas\BasicStructure.DocumentAsData.g.cs;Schemas\BasicStructure.docx" Condition="'{{includeGeneratedSourceAsContent}}' == 'True'" />
                <DocumentAsData Include="Schemas\BasicStructure.docx" Condition="'{{includeDocument}}' == 'True'" />
              </ItemGroup>
              <Import Project="{{SecurityElement.Escape(TargetsPath)}}" />
              <Target Name="WriteCompileItems">
                <!-- AvailableItemNameは、IDEへ公開するビルドアクションの候補です。 -->
                <WriteLinesToFile File="AvailableItemNames.txt" Lines="@(AvailableItemName)" Overwrite="true" />
                <WriteLinesToFile File="None.txt" Lines="@(None->'%(FullPath)')" Overwrite="true" />
                <WriteLinesToFile File="Content.txt" Lines="@(Content->'%(FullPath)')" Overwrite="true" />
                <WriteLinesToFile File="Compile.txt" Lines="@(Compile->'%(FullPath)')" Overwrite="true" />
                <!-- DependentUponは、生成ソースを元文書の子として表示するための親ファイル名です。 -->
                <WriteLinesToFile File="CompileNesting.txt" Lines="@(Compile->'%(FullPath)|%(DependentUpon)')" Overwrite="true" />
                <!-- LastGenOutputは、元文書から対応する生成ソースを識別するためのファイル名です。 -->
                <WriteLinesToFile File="DocumentGeneratedOutput.txt" Lines="@(DocumentAsData->'%(FullPath)|%(LastGenOutput)')" Overwrite="true" />
              </Target>
            </Project>
            """);
        return WritePowerShellScript(
            "SdkProject.ps1",
            """
            dotnet msbuild ./DocumentAsData.SdkProject.csproj /t:WriteCompileItems /nologo /v:minimal
            exit $LASTEXITCODE
            """);
    }

    /// <summary>
    /// 共通のエラー設定とUTF-8出力設定を付けて、検証用スクリプトを書き込みます。
    /// 終了コードの扱いは、呼び出し元のコマンド列に残します。
    /// </summary>
    /// <param name="fileName">一時プロジェクト内に作成するファイル名。</param>
    /// <param name="commands">サンプル固有のコマンドと終了処理。</param>
    /// <returns>作成したスクリプトのパス。</returns>
    string WritePowerShellScript(string fileName, string commands)
    {
        var scriptFilePath = Path.Combine(DirectoryPath, fileName);
        File.WriteAllText(
            scriptFilePath,
            $$"""
            $ErrorActionPreference = 'Stop'
            [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            {{commands}}
            """);
        return scriptFilePath;
    }

    /// <summary>
    /// 新しい一時プロジェクトを作成します。
    /// </summary>
    /// <returns>作成した一時プロジェクト。</returns>
    internal static MSBuild連携テストプロジェクト Create() =>
        new();

    /// <summary>
    /// テスト中に作成した一時プロジェクトを削除します。
    /// </summary>
    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
        {
            Directory.Delete(DirectoryPath, true);
        }
    }

    /// <summary>
    /// 既存の基本構造テスト文書を、指定した相対パスへ追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>追加したWord文書の絶対パス。</returns>
    internal string AddBasicStructureDocument(string relativePath) =>
        AddDocument(BasicStructureDocumentFilePath, relativePath);

    /// <summary>
    /// 既存の顧客データテスト文書を、指定した相対パスへ追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>追加したWord文書の絶対パス。</returns>
    internal string AddCustomerDataDocument(string relativePath) =>
        AddDocument(CustomerDataDocumentFilePath, relativePath);

    /// <summary>
    /// customer_idとcustomer-idが同じ生成名になる既存の文書を追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>追加したWord文書の絶対パス。</returns>
    internal string AddMergeFieldNameCollisionDocument(string relativePath) =>
        AddDocument(@"TestData\コード生成\MERGEFIELD名衝突.docx", relativePath);

    /// <summary>
    /// プロジェクト直下へ、指定したWord文書用の識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="documentFileName">対応するWord文書のファイル名。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddProjectDictionaryFor(
        string documentFileName,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                DictionaryFileName(documentFileName)),
            json);

    /// <summary>
    /// Word文書と同じディレクトリへ、識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="documentRelativePath">対応するWord文書の一時プロジェクト内相対パス。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddDocumentDictionaryFor(
        string documentRelativePath,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                Path.GetDirectoryName(documentRelativePath) ?? "",
                DictionaryFileName(documentRelativePath)),
            json);

    /// <summary>
    /// MSBuildタスクを直接実行します。
    /// </summary>
    /// <param name="documentFilePaths">生成対象Word文書の絶対パス。</param>
    /// <returns>タスクの実行結果。</returns>
    internal MSBuild連携タスク実行結果 Generate(params string[] documentFilePaths)
    {
        var buildEngine = new RecordingBuildEngine();
        var task = new GenerateDocumentAsData
        {
            BuildEngine = buildEngine,
            DocumentFiles =
            [
                .. from filePath in documentFilePaths
                   select new TaskItem(filePath)
            ],
            ProjectDirectory = DirectoryPath
        };

        return new(
            task.Execute(),
            task.GeneratedFiles,
            buildEngine.Warnings,
            buildEngine.Errors);
    }

    /// <summary>
    /// 指定したWord文書に対応する生成ファイルの絶対パスを取得します。
    /// </summary>
    /// <param name="documentRelativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>生成ファイルの絶対パス。</returns>
    internal string GeneratedFilePathFor(string documentRelativePath) =>
        Path.Combine(
            DirectoryPath,
            Path.GetDirectoryName(documentRelativePath) ?? "",
            $"{Path.GetFileNameWithoutExtension(documentRelativePath)}.DocumentAsData.g.cs");

    /// <summary>
    /// 既存のWord文書を一時プロジェクトへコピーします。
    /// </summary>
    /// <param name="sourcePath">コピー元のWord文書。</param>
    /// <param name="relativePath">一時プロジェクト内のWord文書相対パス。</param>
    /// <returns>コピーしたWord文書の絶対パス。</returns>
    string AddDocument(
        string sourcePath,
        string relativePath)
    {
        var destinationPath = Path.Combine(DirectoryPath, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        File.Copy(sourcePath, destinationPath);
        return destinationPath;
    }

    /// <summary>
    /// 識別子名変換辞書を書き込みます。
    /// </summary>
    /// <param name="filePath">書き込み先ファイル。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    static string WriteDictionary(
        string filePath,
        string json)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
        File.WriteAllText(filePath, json);
        return filePath;
    }

    /// <summary>
    /// Word文書名に対応する識別子名変換辞書のファイル名を作ります。
    /// </summary>
    /// <param name="documentFileName">対応するWord文書のファイル名。</param>
    /// <returns>識別子名変換辞書のファイル名。</returns>
    static string DictionaryFileName(string documentFileName) =>
        $"{Path.GetFileNameWithoutExtension(documentFileName)}.documentasdata.json";
}

/// <summary>
/// MSBuildタスクの実行結果を、テスト内で観測しやすい形にまとめます。
/// </summary>
/// <param name="Succeeded">タスクが成功した場合は <see langword="true"/>。</param>
/// <param name="GeneratedFiles">タスクが返した生成ファイル。</param>
/// <param name="Warnings">タスクが記録した警告。</param>
/// <param name="Errors">タスクが記録したエラー。</param>
sealed record MSBuild連携タスク実行結果(
    bool Succeeded,
    ITaskItem[] GeneratedFiles,
    IReadOnlyList<BuildWarningEventArgs> Warnings,
    IReadOnlyList<BuildErrorEventArgs> Errors)
{
    /// <summary>
    /// 生成ファイルの絶対パス一覧です。
    /// </summary>
    internal string[] GeneratedFilePaths =>
        [.. GeneratedFiles.Select(it => it.ItemSpec)];

    /// <summary>
    /// 生成ファイルが一つであるテストで、そのMSBuild項目を取得します。
    /// </summary>
    internal ITaskItem SingleGeneratedFile =>
        GeneratedFiles.Single();

    /// <summary>
    /// 生成ファイルが一つであるテストで、そのファイルパスを取得します。
    /// </summary>
    internal string SingleGeneratedFilePath =>
        GeneratedFilePaths.Single();

    /// <summary>
    /// 生成ファイルが一つであるテストで、そのC#ソースを取得します。
    /// </summary>
    internal string SingleGeneratedSource =>
        File.ReadAllText(SingleGeneratedFilePath);
}

/// <summary>
/// PowerShellの終了コードと診断出力を保持します。
/// </summary>
/// <param name="ExitCode">プロセスの終了コード。</param>
/// <param name="Output">標準出力と標準エラー。</param>
sealed record PowerShell実行結果(int ExitCode, string Output)
{
    /// <summary>
    /// スクリプトを別プロセスで実行し、両方の出力を同時に読み出してバッファ待ちを防ぎます。
    /// </summary>
    /// <param name="scriptFilePath">実行するスクリプト。</param>
    /// <param name="workingDirectory">MSBuildの作業ディレクトリ。</param>
    /// <returns>終了コードと出力。</returns>
    internal static PowerShell実行結果 Run(string scriptFilePath, string workingDirectory)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "pwsh",
                ArgumentList = { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFilePath },
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };
        process.Start();
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        process.WaitForExit();
        return new(process.ExitCode, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
    }
}

/// <summary>
/// MSBuildタスクが出力した警告とエラーを記録します。
/// </summary>
sealed class RecordingBuildEngine : IBuildEngine
{
    /// <summary>
    /// タスクが記録した警告です。
    /// </summary>
    internal List<BuildWarningEventArgs> Warnings { get; } = [];

    /// <summary>
    /// タスクが記録したエラーです。
    /// </summary>
    internal List<BuildErrorEventArgs> Errors { get; } = [];

    /// <summary>
    /// タスクエラー後も処理を継続するかどうかを取得します。
    /// </summary>
    public bool ContinueOnError => false;

    /// <summary>
    /// テスト用のタスク要素行番号を取得します。
    /// </summary>
    public int LineNumberOfTaskNode => 0;

    /// <summary>
    /// テスト用のタスク要素列番号を取得します。
    /// </summary>
    public int ColumnNumberOfTaskNode => 0;

    /// <summary>
    /// テスト用のプロジェクトファイル名を取得します。
    /// </summary>
    public string ProjectFileOfTaskNode => "";

    /// <summary>
    /// このテスト補助では別のMSBuildプロジェクトをビルドしません。
    /// </summary>
    public bool BuildProjectFile(
        string projectFileName,
        string[] targetNames,
        IDictionary globalProperties,
        IDictionary targetOutputs) =>
        throw new NotImplementedException();

    /// <summary>
    /// カスタムイベントを受け取ります。
    /// </summary>
    public void LogCustomEvent(CustomBuildEventArgs e)
    {
    }

    /// <summary>
    /// エラーイベントを記録します。
    /// </summary>
    public void LogErrorEvent(BuildErrorEventArgs e)
    {
        Errors.Add(e);
    }

    /// <summary>
    /// メッセージイベントを受け取ります。
    /// </summary>
    public void LogMessageEvent(BuildMessageEventArgs e)
    {
    }

    /// <summary>
    /// 警告イベントを記録します。
    /// </summary>
    public void LogWarningEvent(BuildWarningEventArgs e)
    {
        Warnings.Add(e);
    }
}
