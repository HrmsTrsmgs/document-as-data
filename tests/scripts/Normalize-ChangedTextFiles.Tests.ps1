$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '../../scripts/Normalize-ChangedTextFiles.ps1'
$startInfo = [System.Diagnostics.ProcessStartInfo]::new()
$startInfo.FileName = 'pwsh'
$startInfo.ArgumentList.Add('-NoProfile')
$startInfo.ArgumentList.Add('-File')
$startInfo.ArgumentList.Add($scriptPath)
$startInfo.ArgumentList.Add('-Check')
$startInfo.ArgumentList.Add('-IncludeUntracked')
$startInfo.UseShellExecute = $false
$startInfo.RedirectStandardOutput = $true
$startInfo.RedirectStandardError = $true
$startInfo.Environment['GIT_DIR'] = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())

$process = [System.Diagnostics.Process]::Start($startInfo)
try {
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    if ($process.ExitCode -eq 0) {
        throw "Git failed to obtain changed files, but the check succeeded. stdout: $($stdoutTask.Result) stderr: $($stderrTask.Result)"
    }
}
finally {
    $process.Dispose()
}
