#Requires -Version 7.0
<#
.SYNOPSIS
Release構成を検証し、公開せずにNuGetパッケージと検証結果を保存します。
.PARAMETER VisualStudioMSBuild
追加検証に使用するVisual StudioのMSBuild.exe。省略時はdotnet版だけを検証します。
#>
[CmdletBinding()]
param([string]$VisualStudioMSBuild)

$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$runDirectory = Join-Path $repositoryRoot ('artifacts/release/' + [Guid]::NewGuid().ToString('N'))
$packageDirectory = Join-Path $runDirectory 'packages'
$previousTestMSBuild = $env:DOCUMENTASDATA_TEST_MSBUILD

function Invoke-DotNet {
    # ネイティブコマンドの非ゼロ終了は、PowerShellのErrorActionPreferenceだけでは停止しません。
    param([string[]]$CommandArguments)
    & dotnet @CommandArguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet $($CommandArguments -join ' ') failed with exit code $LASTEXITCODE."
    }
}

Push-Location $repositoryRoot
try {
    if ($VisualStudioMSBuild) {
        $VisualStudioMSBuild = (Resolve-Path -LiteralPath $VisualStudioMSBuild).Path
    }
    # 前の検証結果と混ぜず、新しいディレクトリへ出力します。既存成果物は削除しません。
    New-Item -ItemType Directory -Path $packageDirectory | Out-Null
    $env:DOCUMENTASDATA_TEST_MSBUILD = $null
    Invoke-DotNet @('restore', './DocumentAsData.slnx', '-p:TreatWarningsAsErrors=true')
    Invoke-DotNet @(
        'build', './DocumentAsData.slnx', '--no-restore', '--configuration', 'Release',
        '--disable-build-servers', '-p:UseSharedCompilation=false', '-p:TreatWarningsAsErrors=true'
    )
    Invoke-DotNet @(
        'test', './DocumentAsData.slnx', '--no-build', '--no-restore', '--configuration', 'Release',
        '--logger', 'trx', '--results-directory', (Join-Path $runDirectory 'tests')
    )
    Invoke-DotNet @('format', './DocumentAsData.slnx', '--no-restore', '--verify-no-changes', '--severity', 'warn')

    if ($VisualStudioMSBuild) {
        $env:DOCUMENTASDATA_TEST_MSBUILD = $VisualStudioMSBuild
        Invoke-DotNet @(
            'test', './tests/DocumentAsData.CodeGeneration.Tests', '--no-build', '--no-restore',
            '--configuration', 'Release', '--filter', 'FullyQualifiedName~パッケージ',
            '--logger', 'trx', '--results-directory', (Join-Path $runDirectory 'visual-studio-tests')
        )
        $env:DOCUMENTASDATA_TEST_MSBUILD = $null
    }

    foreach ($projectName in @('DocumentAsData', 'DocumentAsData.CodeGeneration', 'DocumentAsData.Build', 'DocumentAsData.Package')) {
        Invoke-DotNet @(
            'pack', "./src/$projectName/$projectName.csproj", '--no-build', '--no-restore',
            '--configuration', 'Release', '--disable-build-servers', '-p:TreatWarningsAsErrors=true',
            '--output', $packageDirectory
        )
    }

    # 公開時に同じ成果物を選べるよう、ハッシュと検証対象コミットを添えます。
    Get-ChildItem -LiteralPath $packageDirectory -Filter '*.nupkg' |
        Sort-Object Name |
        ForEach-Object { '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash, $_.Name } |
        Set-Content -LiteralPath (Join-Path $runDirectory 'SHA256SUMS.txt') -Encoding utf8
    $verifiedCommit = & git -c "safe.directory=$($repositoryRoot.Replace('\', '/'))" rev-parse HEAD
    if ($LASTEXITCODE -ne 0) { throw '検証対象コミットを取得できませんでした。' }
    $workingChanges = & git -c "safe.directory=$($repositoryRoot.Replace('\', '/'))" status --porcelain
    if ($LASTEXITCODE -ne 0) { throw '作業ツリーの状態を取得できませんでした。' }
    @(
        '# ローカルRelease検証結果'
        ''
        "- コミット: $verifiedCommit"
        "- 作業ツリー変更あり: $([bool]$workingChanges)"
        "- 検証完了日時（UTC）: $([DateTime]::UtcNow.ToString('O'))"
        "- .NET SDK: $(& dotnet --version)"
        '- Releaseビルド・全テスト・書式検査・警告をエラーにした梱包: 成功'
        "- Visual Studio版MSBuildの追加検証: $([bool]$VisualStudioMSBuild)"
        '- 詳細な件数・結果: tests/*.trx（追加検証時はvisual-studio-tests/*.trxも参照）'
        '- GitHubへのpush、タグ作成、Release作成、NuGet公開: 未実行'
    ) | Set-Content -LiteralPath (Join-Path $runDirectory 'VERIFICATION.md') -Encoding utf8
    Write-Host "公開前の検証と梱包が完了しました: $runDirectory"
}
finally {
    $env:DOCUMENTASDATA_TEST_MSBUILD = $previousTestMSBuild
    Pop-Location
}
