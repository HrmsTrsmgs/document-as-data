param(
    [switch]$Check,
    [switch]$IncludeUntracked
)

$ErrorActionPreference = "Stop"

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$safeDirectory = $repositoryRoot.Replace("\", "/")

function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

    & git -C $repositoryRoot -c "safe.directory=$safeDirectory" -c "core.quotePath=false" @Arguments
}

function Get-ChangedPath {
    $paths = @(
        Invoke-Git "diff" "--name-only" "--diff-filter=ACMRT"
        Invoke-Git "diff" "--cached" "--name-only" "--diff-filter=ACMRT"
    )

    if ($IncludeUntracked) {
        $paths += Invoke-Git "ls-files" "--others" "--exclude-standard"
    }

    $paths |
        Where-Object { $_ } |
        Sort-Object -Unique
}

function Get-TextFilePolicy {
    param([string]$Path)

    $fileName = [IO.Path]::GetFileName($Path)
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($fileName -in @(
            ".editorconfig",
            ".gitattributes",
            ".gitignore"
        )) {
        return @{
            HasBom = $false
        }
    }

    if ($extension -eq ".cs") {
        return @{
            HasBom = $true
        }
    }

    if ($extension -in @(
            ".config",
            ".csproj",
            ".json",
            ".md",
            ".props",
            ".ps1",
            ".rb",
            ".sln",
            ".slnx",
            ".targets",
            ".txt",
            ".xml",
            ".yaml",
            ".yml"
        )) {
        return @{
            HasBom = $false
        }
    }

    return $null
}

function ConvertTo-NormalizedText {
    param([string]$Text)

    $lines = $Text -split "\r\n|\n|\r"

    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq "") {
        if ($lines.Count -eq 1) {
            $lines = @()
        }
        else {
            $lines = $lines[0..($lines.Count - 2)]
        }
    }

    if ($lines.Count -eq 0) {
        return ""
    }

    $normalizedLines =
        foreach ($line in $lines) {
            $line.TrimEnd()
        }

    return ($normalizedLines -join "`r`n") + "`r`n"
}

$changedPaths = Get-ChangedPath
$failed = $false

foreach ($changedPath in $changedPaths) {
    $path = Join-Path $repositoryRoot $changedPath

    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        continue
    }

    $policy = Get-TextFilePolicy $changedPath

    if ($null -eq $policy) {
        continue
    }

    $currentBytes = [IO.File]::ReadAllBytes($path)
    $utf8 = [Text.UTF8Encoding]::new($false, $true)
    $text = $utf8.GetString($currentBytes)
    $normalizedText = ConvertTo-NormalizedText $text
    $encoding = [Text.UTF8Encoding]::new([bool]$policy.HasBom)
    $normalizedBytes = $encoding.GetBytes($normalizedText)

    if (-not [Linq.Enumerable]::SequenceEqual($currentBytes, $normalizedBytes)) {
        if ($Check) {
            Write-Host "Needs normalization: $changedPath"
            $failed = $true
        }
        else {
            [IO.File]::WriteAllBytes($path, $normalizedBytes)
            Write-Host "Normalized: $changedPath"
        }
    }
}

if ($failed) {
    exit 1
}
