param(
    [Parameter(Mandatory)][string]$VSTestConsole,
    [Parameter(Mandatory)][string]$TestAssembly
)

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '../../scripts/Prepare-Release.ps1'
$tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())

function Get-VisualStudioTestFilter([string]$ReleaseScript, [string]$MSBuildPath) {
    # この関数のスコープでだけ dotnet を置き換え、追加検証のフィルターを取得します。
    $LASTEXITCODE = 0
    function dotnet {
        $filterIndex = [array]::IndexOf($args, '--filter')
        if ($filterIndex -ge 0) {
            Write-Error -Message 'Captured Visual Studio test filter.' -ErrorId 'CapturedFilter' -TargetObject $args[$filterIndex + 1] -ErrorAction Stop
        }
    }
    try {
        & $ReleaseScript -VisualStudioMSBuild $MSBuildPath
    }
    catch {
        if ($_.FullyQualifiedErrorId -ne 'CapturedFilter,dotnet') { throw }
        return $_.TargetObject
    }
    throw 'Visual Studio validation was not invoked.'
}

function Get-SelectedTests([string]$Filter) {
    $output = @(& dotnet $VSTestConsole $TestAssembly /ListTests "/TestCaseFilter:$Filter" 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "VSTest discovery failed: $output" }
    @($output | Where-Object { $_ -match '^\s+Marimo\.DocumentAsData\.' } | ForEach-Object { $_.Trim() })
}

try {
    $scriptDirectory = Join-Path $tempDirectory 'scripts'
    New-Item -ItemType Directory -Path $scriptDirectory | Out-Null
    $releaseScript = Copy-Item -LiteralPath $scriptPath -Destination $scriptDirectory -PassThru
    $visualStudioMSBuild = Join-Path $tempDirectory 'MSBuild.exe'
    New-Item -ItemType File -Path $visualStudioMSBuild | Out-Null

    $filter = Get-VisualStudioTestFilter $releaseScript.FullName $visualStudioMSBuild
    $previousSelection = @(Get-SelectedTests 'FullyQualifiedName~パッケージ')
    if ($previousSelection.Count -eq 0) { throw 'VSTest did not discover the existing package tests.' }
    $selection = @(Get-SelectedTests $filter)
    $required = @(
        $previousSelection
        'Marimo.DocumentAsData.CodeGeneration.Test.MSBuild連携タスクのテスト.共用ライブラリのBuild参照を非公開にしてもCoreは利用アプリへ渡ります'
        'Marimo.DocumentAsData.CodeGeneration.Test.MSBuild連携タスクのテスト.共用ライブラリの利用アプリにはCoreだけが渡り生成用DLLは配布されません'
    )
    $missing = @($required | Where-Object { $_ -notin $selection })
    if ($missing.Count -gt 0) { throw "Visual Studio validation omitted tests: $($missing -join ', ')" }
}
finally {
    if (Test-Path -LiteralPath $tempDirectory) { Remove-Item -LiteralPath $tempDirectory -Recurse -Force }
}
