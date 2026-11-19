param(
    [Parameter(Mandatory)][string]$VSTestConsole,
    [Parameter(Mandatory)][string]$TestAssembly
)

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '../../scripts/Prepare-Release.ps1'
$tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())

try {
    $scriptDirectory = Join-Path $tempDirectory 'scripts'
    New-Item -ItemType Directory -Path $scriptDirectory | Out-Null
    Copy-Item -LiteralPath $scriptPath -Destination $scriptDirectory
    $visualStudioMSBuild = Join-Path $tempDirectory 'MSBuild.exe'
    New-Item -ItemType File -Path $visualStudioMSBuild | Out-Null

    # 追加検証の引数を受け取るまで、実際のビルドや梱包を行いません。
    function dotnet {
        $global:LASTEXITCODE = 0
        if ($args -contains '--filter') {
            $global:releaseTestFilter = $args[[array]::IndexOf($args, '--filter') + 1]
            throw 'CapturedVisualStudioFilter'
        }
    }
    try {
        & (Join-Path $scriptDirectory 'Prepare-Release.ps1') -VisualStudioMSBuild $visualStudioMSBuild
        throw 'Visual Studio validation was not invoked.'
    }
    catch {
        if ($_.Exception.Message -ne 'CapturedVisualStudioFilter') { throw }
    }
    Remove-Item function:dotnet

    function Get-SelectedTests([string]$Filter) {
        $output = @(& dotnet $VSTestConsole $TestAssembly /ListTests "/TestCaseFilter:$Filter" 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "VSTest discovery failed: $output" }
        @($output | Where-Object { $_ -match '^\s+Marimo\.DocumentAsData\.' } | ForEach-Object { $_.Trim() })
    }

    $previousSelection = @(Get-SelectedTests 'FullyQualifiedName~パッケージ')
    if ($previousSelection.Count -eq 0) { throw 'VSTest did not discover the existing package tests.' }
    $selection = @(Get-SelectedTests $global:releaseTestFilter)
    $required = @(
        $previousSelection
        'Marimo.DocumentAsData.CodeGeneration.Test.MSBuild連携タスクのテスト.共用ライブラリのBuild参照を非公開にしてもCoreは利用アプリへ渡ります'
        'Marimo.DocumentAsData.CodeGeneration.Test.MSBuild連携タスクのテスト.共用ライブラリの利用アプリにはCoreだけが渡り生成用DLLは配布されません'
    )
    $missing = @($required | Where-Object { $_ -notin $selection })
    if ($missing.Count -gt 0) { throw "Visual Studio validation omitted tests: $($missing -join ', ')" }
}
finally {
    Remove-Item function:dotnet -ErrorAction SilentlyContinue
    Remove-Variable releaseTestFilter -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $tempDirectory) { Remove-Item -LiteralPath $tempDirectory -Recurse -Force }
}
