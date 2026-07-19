param(
    [string]$InputPath = "demo\.working\ROneCOne_Delegates_Demo.xlsx",
    [string]$OutputPath = "demo\ROneCOne_Delegates_Demo.xlsm",
    [ValidateRange(5, 120)]
    [int]$TimeoutSeconds = 30,
    [switch]$Worker,
    [string]$ProcessInfoPath = "demo\.working\convert-process.json"
)

$ErrorActionPreference = "Stop"

$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

$resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$resolvedProcessInfo = [System.IO.Path]::GetFullPath($ProcessInfoPath)
$workingDirectory = Split-Path -Parent $resolvedProcessInfo
$dialogLog = Join-Path $workingDirectory "convert-dialogs.jsonl"
$watcherStop = Join-Path $workingDirectory "convert-watcher.stop"

if ($Worker) {
    $excel = $null
    $workbook = $null
    $watcher = $null
    try {
        [System.IO.Directory]::CreateDirectory($workingDirectory) | Out-Null
        Remove-Item -LiteralPath $dialogLog, $watcherStop -Force -ErrorAction SilentlyContinue
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3

        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ROneCOneConvertProcess
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
        [uint32]$excelProcessId = 0
        [void][ROneCOneConvertProcess]::GetWindowThreadProcessId(
            [IntPtr]$excel.Hwnd,
            [ref]$excelProcessId)

        $watcherScript = Join-Path $PSScriptRoot "watch_vbe_dialogs.ps1"
        $watcherArguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$watcherScript`"",
            "-ExcelProcessId", $excelProcessId,
            "-LogPath", "`"$dialogLog`"",
            "-StopPath", "`"$watcherStop`"",
            "-TimeoutSeconds", 120,
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
            excel_process_id = $excelProcessId
            watcher_process_id = $watcher.Id
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $resolvedProcessInfo

        $workbook = $excel.Workbooks.Open($resolvedInput, 0, $false)
        $workbook.SaveAs($resolvedOutput, 52)
        $bootstrap = $workbook.VBProject.VBComponents.Add(1)
        $bootstrap.Name = "Module1"
        $bootstrap.CodeModule.AddFromString("Option Explicit`r`n")
        $workbook.Save()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrap)
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        $workbook = $null
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if (Test-Path -LiteralPath $dialogLog) {
            $modalRecords = @(
                Get-Content -LiteralPath $dialogLog |
                    ConvertFrom-Json |
                    Where-Object {
                        $_.class_name -eq "#32770" -or $_.dismissal_action -ne "none"
                    }
            )
            if ($modalRecords.Count -gt 0) {
                throw "Office or VBE popup was observed while converting the demo."
            }
        }
        Write-Output $resolvedOutput
        exit 0
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
    }
}

$outputDirectory = Split-Path -Parent $resolvedOutput
[System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory((Split-Path -Parent $resolvedProcessInfo)) | Out-Null
$stdoutPath = [System.IO.Path]::ChangeExtension($resolvedProcessInfo, ".stdout.log")
$stderrPath = [System.IO.Path]::ChangeExtension($resolvedProcessInfo, ".stderr.log")
Remove-Item -LiteralPath $resolvedProcessInfo, $stdoutPath, $stderrPath `
    -Force -ErrorAction SilentlyContinue

$arguments = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$PSCommandPath`"",
    "-InputPath", "`"$resolvedInput`"",
    "-OutputPath", "`"$resolvedOutput`"",
    "-ProcessInfoPath", "`"$resolvedProcessInfo`"",
    "-Worker"
)
$process = Start-Process `
    -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ArgumentList $arguments `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutPath `
    -RedirectStandardError $stderrPath `
    -PassThru

try {
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        if (Test-Path -LiteralPath $resolvedProcessInfo) {
            try {
                $owned = Get-Content -Raw -LiteralPath $resolvedProcessInfo | ConvertFrom-Json
                Stop-Process -Id $owned.excel_process_id -Force -ErrorAction SilentlyContinue
                Stop-Process -Id $owned.watcher_process_id -Force -ErrorAction SilentlyContinue
            }
            catch {}
        }
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        throw "Workbook conversion exceeded the hard $TimeoutSeconds-second deadline."
    }
    $process.WaitForExit()
    if ((Test-Path -LiteralPath $stderrPath) -and
        (Get-Item -LiteralPath $stderrPath).Length -gt 0) {
        Get-Content -LiteralPath $stderrPath | ForEach-Object {
            [Console]::Error.WriteLine($_)
        }
        throw "Workbook conversion worker failed."
    }
    Get-Content -LiteralPath $stdoutPath | Write-Output
}
finally {
    if (-not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    $process.Dispose()
}
