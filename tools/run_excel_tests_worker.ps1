param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$ProcessInfoPath,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0.01, 60)]
    [double]$MaxBenchmarkSeconds,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0.01, 60)]
    [double]$MaxCollectionBenchmarkSeconds,

    [Parameter(Mandatory = $true)]
    [ValidateRange(0.01, 60)]
    [double]$MaxOrderingBenchmarkSeconds
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

    $collectionSheet = $null
    try {
        $collectionSheet = $workbook.Worksheets.Item("Collection Results")
    }
    catch {
        $collectionSheet = $workbook.Worksheets.Add()
        $collectionSheet.Name = "Collection Results"
    }

    $collectionBenchmarkSheet = $null
    try {
        $collectionBenchmarkSheet = $workbook.Worksheets.Item("Collection Benchmarks")
    }
    catch {
        $collectionBenchmarkSheet = $workbook.Worksheets.Add()
        $collectionBenchmarkSheet.Name = "Collection Benchmarks"
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
    $collectionSheet.Range("A1:B5").ClearContents()
    $collectionSheet.Range("A1").Value2 = "ROneCOne generic collection test suite"
    $collectionSheet.Range("A2").Value2 = "Passed"
    $collectionSheet.Range("A3").Value2 = "Failed"
    $collectionSheet.Range("A4").Value2 = "Status"
    $collectionSheet.Range("A5").Value2 = "Error"
    $collectionBenchmarkSheet.Range("A1:B8").ClearContents()
    $collectionBenchmarkSheet.Range("A1").Value2 = "ROneCOne collection benchmark"
    $collectionBenchmarkSheet.Range("A2").Value2 = "Source elements"
    $collectionBenchmarkSheet.Range("A3").Value2 = "Seconds"
    $collectionBenchmarkSheet.Range("A4").Value2 = "Filtered elements"
    $collectionBenchmarkSheet.Range("A5").Value2 = "Ordering source elements"
    $collectionBenchmarkSheet.Range("A6").Value2 = "Composite ordering seconds"
    $collectionBenchmarkSheet.Range("A7").Value2 = "Ordered elements"
    $collectionBenchmarkSheet.Range("A8").Value2 = "First ordered result"
    $macroPrefix = "'" + $workbook.Name.Replace("'", "''") + "'!"
    $stage = "run delegate tests"
    $excel.Run($macroPrefix + "RunROneCOneTests") | Out-Null
    $stage = "run generic collection tests"
    $excel.Run($macroPrefix + "RunROneCOneCollectionTests") | Out-Null
    $stage = "run delegate benchmark"
    $excel.Run($macroPrefix + "RunROneCOneBenchmark") | Out-Null
    $stage = "run generic collection benchmark"
    $excel.Run($macroPrefix + "RunROneCOneCollectionBenchmark") | Out-Null
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
    $collectionStatus = [string]$collectionSheet.Range("B4").Value2
    $collectionPassed = [int]$collectionSheet.Range("B2").Value2
    $collectionFailed = [int]$collectionSheet.Range("B3").Value2
    $seconds = [double]$benchmarkSheet.Range("B3").Value2
    $collectionSeconds = [double]$collectionBenchmarkSheet.Range("B3").Value2
    $orderingSeconds = [double]$collectionBenchmarkSheet.Range("B6").Value2

    $stage = "validate observed results"
    if ($status -ne "PASS" -or $failed -ne 0) {
        $errorDetail = [string]$testSheet.Range("B5").Value2
        if ([string]::IsNullOrWhiteSpace($errorDetail)) {
            $failedAssertions = @()
            for ($row = 6; $row -le 200; $row++) {
                if ([string]$testSheet.Cells.Item($row, 2).Value2 -eq "FAIL") {
                    $failedAssertions += "{0}: {1}" -f `
                        [string]$testSheet.Cells.Item($row, 1).Value2,
                        [string]$testSheet.Cells.Item($row, 3).Value2
                }
            }
            $errorDetail = $failedAssertions -join "; "
        }
        throw "Live Excel tests failed: status=$status passed=$passed failed=$failed detail=$errorDetail"
    }
    if ($collectionStatus -ne "PASS" -or $collectionFailed -ne 0) {
        $errorDetail = [string]$collectionSheet.Range("B5").Value2
        if ([string]::IsNullOrWhiteSpace($errorDetail)) {
            $failedAssertions = @()
            for ($row = 6; $row -le 300; $row++) {
                if ([string]$collectionSheet.Cells.Item($row, 2).Value2 -eq "FAIL") {
                    $failedAssertions += "{0}: {1}" -f `
                        [string]$collectionSheet.Cells.Item($row, 1).Value2,
                        [string]$collectionSheet.Cells.Item($row, 3).Value2
                }
            }
            $errorDetail = $failedAssertions -join "; "
        }
        throw "Collection tests failed: status=$collectionStatus passed=$collectionPassed failed=$collectionFailed detail=$errorDetail"
    }
    if ($seconds -le 0 -or $seconds -gt $MaxBenchmarkSeconds) {
        throw "Delegate benchmark missed gate: seconds=$seconds maximum=$MaxBenchmarkSeconds"
    }
    if ($collectionSeconds -le 0 -or `
        $collectionSeconds -gt $MaxCollectionBenchmarkSeconds) {
        throw "Collection benchmark missed gate: seconds=$collectionSeconds maximum=$MaxCollectionBenchmarkSeconds"
    }
    if ($orderingSeconds -le 0 -or `
        $orderingSeconds -gt $MaxOrderingBenchmarkSeconds) {
        throw "Ordering benchmark missed gate: seconds=$orderingSeconds maximum=$MaxOrderingBenchmarkSeconds"
    }
    if ([int]$collectionBenchmarkSheet.Range("B7").Value2 -ne 10000 -or `
        [int]$collectionBenchmarkSheet.Range("B8").Value2 -ne 10000) {
        throw "Ordering benchmark produced an invalid composite order."
    }

    [pscustomobject]@{
        workbook = $resolvedWorkbook
        status = $status
        passed = $passed
        failed = $failed
        collection_status = $collectionStatus
        collection_passed = $collectionPassed
        collection_failed = $collectionFailed
        benchmark_invocations = [int]$benchmarkSheet.Range("B2").Value2
        benchmark_seconds = $seconds
        benchmark_gate_seconds = $MaxBenchmarkSeconds
        collection_benchmark_elements = [int]$collectionBenchmarkSheet.Range("B2").Value2
        collection_benchmark_seconds = $collectionSeconds
        collection_benchmark_gate_seconds = $MaxCollectionBenchmarkSeconds
        ordering_benchmark_elements = [int]$collectionBenchmarkSheet.Range("B5").Value2
        ordering_benchmark_seconds = $orderingSeconds
        ordering_benchmark_gate_seconds = $MaxOrderingBenchmarkSeconds
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
