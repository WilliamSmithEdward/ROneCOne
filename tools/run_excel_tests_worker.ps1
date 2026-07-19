param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$ProcessInfoPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0.01, 60)]
    [double]$MaxBenchmarkSeconds
)

$ErrorActionPreference = "Stop"
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedProcessInfo = [System.IO.Path]::GetFullPath($ProcessInfoPath)
$excel = $null
$workbook = $null
$watcher = $null
$excelProcessId = 0
$watcherStop = Join-Path $PSScriptRoot "..\tests\output\vbe-watcher.stop"
$watcherLog = Join-Path $PSScriptRoot "..\tests\output\vbe-dialogs.jsonl"
$stage = "create Excel application"

function Write-ProcessOwnership {
    [ordered]@{
        worker_process_id = $PID
        excel_process_id = $excelProcessId
        watcher_process_id = if ($null -eq $watcher) { 0 } else { $watcher.Id }
    } | ConvertTo-Json -Compress | Set-Content -LiteralPath $resolvedProcessInfo
}

try {
    Remove-Item -LiteralPath $watcherStop -Force -ErrorAction SilentlyContinue

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ROneCOneExcelProcess
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
    [uint32]$ownedExcelProcessId = 0
    [void][ROneCOneExcelProcess]::GetWindowThreadProcessId(
        [IntPtr]$excel.Hwnd,
        [ref]$ownedExcelProcessId)
    $excelProcessId = [int]$ownedExcelProcessId
    Write-ProcessOwnership

    $watcherScript = Join-Path $PSScriptRoot "watch_vbe_dialogs.ps1"
    $watcherArguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$watcherScript`"",
        "-ExcelProcessId", $excelProcessId,
        "-LogPath", "`"$watcherLog`"",
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
    Write-ProcessOwnership

    $stage = "open workbook"
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)

    $stage = "prepare result sheets"
    $testSheet = $null
    try {
        $testSheet = $workbook.Worksheets.Item("Test Results")
    }
    catch {
        $testSheet = $workbook.Worksheets.Item(1)
        $testSheet.Name = "Test Results"
    }

    $benchmarkSheet = $null
    try {
        $benchmarkSheet = $workbook.Worksheets.Item("Benchmarks")
    }
    catch {
        $benchmarkSheet = $workbook.Worksheets.Add()
        $benchmarkSheet.Name = "Benchmarks"
    }

    $testSheet.Range("A1:B4").ClearContents()
    $testSheet.Range("A1").Value2 = "ROneCOne delegate test suite"
    $testSheet.Range("A2").Value2 = "Passed"
    $testSheet.Range("A3").Value2 = "Failed"
    $testSheet.Range("A4").Value2 = "Status"
    $benchmarkSheet.Range("A1:B4").ClearContents()
    $benchmarkSheet.Range("A1").Value2 = "ROneCOne delegate benchmark"
    $benchmarkSheet.Range("A2").Value2 = "Invocations"
    $benchmarkSheet.Range("A3").Value2 = "Seconds"
    $benchmarkSheet.Range("A4").Value2 = "Last result"
    $macroPrefix = "'" + $workbook.Name.Replace("'", "''") + "'!"
    $stage = "run delegate tests"
    $excel.Run($macroPrefix + "RunROneCOneTests") | Out-Null
    $stage = "run delegate benchmark"
    $excel.Run($macroPrefix + "RunROneCOneBenchmark") | Out-Null
    $stage = "read observed results"
    if (Test-Path -LiteralPath $watcherLog) {
        $dialogRecords = @(Get-Content -LiteralPath $watcherLog | ConvertFrom-Json)
        $modalRecords = @($dialogRecords | Where-Object {
            $_.class_name -eq "#32770" -or $_.dismissal_action -ne "none"
        })
        if ($modalRecords.Count -gt 0) {
            $dialogText = $modalRecords[0].child_text -join " | "
            throw "Office or VBE popup observed: $dialogText"
        }
    }
    $status = [string]$testSheet.Range("B4").Value2
    $passed = [int]$testSheet.Range("B2").Value2
    $failed = [int]$testSheet.Range("B3").Value2
    $seconds = [double]$benchmarkSheet.Range("B3").Value2

    $stage = "validate observed results"
    if ($status -ne "PASS" -or $failed -ne 0) {
        $errorDetail = [string]$testSheet.Range("B5").Value2
        throw "Live Excel tests failed: status=$status passed=$passed failed=$failed detail=$errorDetail"
    }
    if ($seconds -le 0 -or $seconds -gt $MaxBenchmarkSeconds) {
        throw "Delegate benchmark missed gate: seconds=$seconds maximum=$MaxBenchmarkSeconds"
    }

    [pscustomobject]@{
        workbook = $resolvedWorkbook
        status = $status
        passed = $passed
        failed = $failed
        benchmark_invocations = [int]$benchmarkSheet.Range("B2").Value2
        benchmark_seconds = $seconds
        benchmark_gate_seconds = $MaxBenchmarkSeconds
    } | ConvertTo-Json -Compress
}
catch {
    [Console]::Error.WriteLine("Live Excel stage '$stage' failed: $($_.Exception.Message)")
    throw
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
        try {
            $workbook.Close($false)
        }
        catch {
            [Console]::Error.WriteLine("Workbook cleanup warning: $($_.Exception.Message)")
        }
        finally {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            [Console]::Error.WriteLine("Excel cleanup warning: $($_.Exception.Message)")
        }
        finally {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
