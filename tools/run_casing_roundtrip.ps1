param(
    [string]$RuntimePath = "src\ROneCOne.cls",
    [string]$WorkbookPath = "tests\output\ROneCOne_Casing.xlsm",
    [string]$ExportDirectory = "tests\output\casing_export",
    [ValidateRange(5, 300)]
    [int]$TimeoutSeconds = 120,
    [switch]$Worker
)

# Issue #6 end to end: build a workbook holding the runtime and a host module,
# open it in a task-owned Excel, export every module through the VBE, and
# require every code token back exactly as written. A declaration anywhere in
# the project that recases a host name shows up here as a changed token.

$ErrorActionPreference = "Stop"

$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

$repository = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repository ".venv\Scripts\python.exe"
$helper = Join-Path $PSScriptRoot "casing_roundtrip.py"
function Resolve-RepositoryPath([string]$path) {
    [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($repository, $path))
}
$resolvedRuntime = Resolve-RepositoryPath $RuntimePath
$resolvedWorkbook = Resolve-RepositoryPath $WorkbookPath
$resolvedExport = Resolve-RepositoryPath $ExportDirectory
$workingDirectory = Split-Path -Parent $resolvedWorkbook
$dialogLog = Join-Path $workingDirectory "casing-dialogs.jsonl"
$watcherStop = Join-Path $workingDirectory "casing-watcher.stop"
$processInfo = Join-Path $workingDirectory "casing-processes.json"

if ($Worker) {
    $excel = $null
    $workbook = $null
    $watcher = $null
    try {
        Remove-Item -LiteralPath $dialogLog, $watcherStop -Force -ErrorAction SilentlyContinue
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3

        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ROneCOneCasingProcess
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
        [uint32]$ownedExcelProcessId = 0
        [void][ROneCOneCasingProcess]::GetWindowThreadProcessId(
            [IntPtr]$excel.Hwnd,
            [ref]$ownedExcelProcessId)

        $watcherArguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$(Join-Path $PSScriptRoot 'watch_vbe_dialogs.ps1')`"",
            "-ExcelProcessId", [int]$ownedExcelProcessId,
            "-LogPath", "`"$dialogLog`"",
            "-StopPath", "`"$watcherStop`"",
            "-TimeoutSeconds", $TimeoutSeconds,
            "-DismissKnownDialogs",
            "-TerminateOnBreakMode"
        )
        $watcher = Start-Process `
            -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList $watcherArguments `
            -WindowStyle Hidden `
            -PassThru
        [ordered]@{
            worker_process_id = $PID
            excel_process_id = [int]$ownedExcelProcessId
            watcher_process_id = $watcher.Id
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $processInfo

        # Macros stay disabled: the project only has to load, and loading is
        # what applies VBA's one spelling per name across every module.
        $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
        Remove-Item -LiteralPath $resolvedExport -Recurse -Force -ErrorAction SilentlyContinue
        [System.IO.Directory]::CreateDirectory($resolvedExport) | Out-Null
        foreach ($component in $workbook.VBProject.VBComponents) {
            $extension = switch ($component.Type) { 1 { ".bas" } 3 { ".frm" } default { ".cls" } }
            $component.Export((Join-Path $resolvedExport ($component.Name + $extension)))
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($component)
        }
    }
    finally {
        Set-Content -LiteralPath $watcherStop -Value "stop"
        if ($null -ne $watcher) {
            if (-not $watcher.WaitForExit(2000)) {
                Stop-Process -Id $watcher.Id -Force -ErrorAction SilentlyContinue
            }
            $watcher.Dispose()
        }
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    exit 0
}

& $python $helper build $resolvedWorkbook --runtime $resolvedRuntime | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Building the casing workbook failed."
}

Remove-Item -LiteralPath $processInfo -Force -ErrorAction SilentlyContinue
$process = Start-Process `
    -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-WorkbookPath", "`"$resolvedWorkbook`"",
        "-ExportDirectory", "`"$resolvedExport`"",
        "-TimeoutSeconds", $TimeoutSeconds,
        "-Worker") `
    -WindowStyle Hidden `
    -PassThru
if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
    if (Test-Path -LiteralPath $processInfo) {
        $owned = Get-Content -Raw -LiteralPath $processInfo | ConvertFrom-Json
        Stop-Process -Id $owned.excel_process_id -Force -ErrorAction SilentlyContinue
        Stop-Process -Id $owned.watcher_process_id -Force -ErrorAction SilentlyContinue
    }
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    throw "The casing round trip timed out after $TimeoutSeconds seconds."
}
if (Test-Path -LiteralPath $processInfo) {
    # Quit normally releases Excel; stop the recorded process only if it is
    # still that Excel, so a reused process ID is never touched.
    $owned = Get-Content -Raw -LiteralPath $processInfo | ConvertFrom-Json
    Get-Process -Id $owned.excel_process_id -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -eq "EXCEL" } |
        Stop-Process -Force -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath $dialogLog) {
    $modal = @(Get-Content -LiteralPath $dialogLog | ConvertFrom-Json |
        Where-Object { $_.class_name -eq "#32770" -or $_.dismissal_action -ne "none" })
    if ($modal.Count -gt 0) {
        throw "An Office or VBE popup was observed during the casing round trip."
    }
}

& $python $helper compare $resolvedExport --runtime $resolvedRuntime
exit $LASTEXITCODE
