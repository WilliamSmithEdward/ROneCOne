param(
    [string]$WorkbookPath = "demo\ROneCOne_Delegates_Demo.xlsm",
    [string]$MacroName = "RunROneCOneDemo",
    [ValidateRange(5, 120)]
    [int]$TimeoutSeconds = 30,
    [ValidateRange(0.01, 60)]
    [double]$MaxOrderingBenchmarkSeconds = 2.5,
    [switch]$Worker,
    [string]$ProcessInfoPath = "demo\.working\demo-processes.json"
)

$ErrorActionPreference = "Stop"

$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedProcessInfo = [System.IO.Path]::GetFullPath($ProcessInfoPath)
$workingDirectory = Split-Path -Parent $resolvedProcessInfo
$dialogLog = Join-Path $workingDirectory "demo-dialogs.jsonl"
$watcherStop = Join-Path $workingDirectory "demo-watcher.stop"

if ($Worker) {
    $excel = $null
    $workbook = $null
    $watcher = $null
    $excelProcessId = 0
    try {
        Remove-Item -LiteralPath $dialogLog, $watcherStop -Force -ErrorAction SilentlyContinue
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 1

        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ROneCOneDemoProcess
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
        [uint32]$ownedExcelProcessId = 0
        [void][ROneCOneDemoProcess]::GetWindowThreadProcessId(
            [IntPtr]$excel.Hwnd,
            [ref]$ownedExcelProcessId)
        $excelProcessId = [int]$ownedExcelProcessId

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

        $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)
        $macroPrefix = "'" + $workbook.Name.Replace("'", "''") + "'!"
        $excel.Run($macroPrefix + $MacroName) | Out-Null
        $excel.Calculate()

        if (Test-Path -LiteralPath $dialogLog) {
            $modalRecords = @(
                Get-Content -LiteralPath $dialogLog |
                    ConvertFrom-Json |
                    Where-Object {
                        $_.class_name -eq "#32770" -or $_.dismissal_action -ne "none"
                    }
            )
            if ($modalRecords.Count -gt 0) {
                throw "Office or VBE popup was observed while running the demo."
            }
        }

        $demoStatus = [string]$workbook.Worksheets.Item("Start Here").Range("B13").Value2
        $exampleSheetNames = @("Examples")
        $userClassSheet = $null
        try {
            $userClassSheet = $workbook.Worksheets.Item("User Class LINQ")
            $exampleSheetNames += "User Class LINQ"
        }
        catch {}
        finally {
            if ($null -ne $userClassSheet) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $userClassSheet)
            }
        }

        $statuses = @()
        foreach ($sheetName in $exampleSheetNames) {
            $examplesSheet = $workbook.Worksheets.Item($sheetName)
            try {
                $exampleCount = [int]$excel.WorksheetFunction.CountA(
                    $examplesSheet.Range("A6:A100"))
                $lastExampleRow = 5 + $exampleCount
                $statuses += @(
                    $examplesSheet.Range("F6:F$lastExampleRow").Value2 |
                        ForEach-Object { [string]$_ }
                )
            }
            finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $examplesSheet)
            }
        }
        $notPassing = @($statuses | Where-Object { $_ -ne "PASS" })
        if ($demoStatus -ne "PASS" -or $notPassing.Count -ne 0) {
            throw "Demo validation failed: status=$demoStatus examples=$($statuses -join ',')"
        }

        $workbook.Save()
        $featureName = [string]$workbook.Worksheets.Item(
            "Start Here").Range("G8").Value2
        $memberDispatchSeconds = [double]$workbook.Worksheets.Item(
            "Benchmarks").Range("C7").Value2
        $orderingSeconds = [double]$workbook.Worksheets.Item(
            "Benchmarks").Range("C8").Value2
        if ($featureName -eq "List<T> + LINQ" -and $memberDispatchSeconds -gt 0.75) {
            throw "Member-dispatch benchmark exceeded the 0.75-second release gate."
        }
        if ($featureName -eq "List<T> + LINQ" -and `
            ($orderingSeconds -le 0 -or `
                $orderingSeconds -gt $MaxOrderingBenchmarkSeconds)) {
            throw "Composite ordering benchmark exceeded the $MaxOrderingBenchmarkSeconds-second gate."
        }
        [pscustomobject]@{
            workbook = $resolvedWorkbook
            feature = $featureName
            status = $demoStatus
            examples_passing = $statuses.Count
            benchmark_seconds = [double]$workbook.Worksheets.Item("Benchmarks").Range("C6").Value2
            member_dispatch_seconds = $memberDispatchSeconds
            ordering_seconds = $orderingSeconds
            ordering_gate_seconds = $MaxOrderingBenchmarkSeconds
        } | ConvertTo-Json -Compress
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

[System.IO.Directory]::CreateDirectory($workingDirectory) | Out-Null
$stdoutPath = Join-Path $workingDirectory "demo-worker.stdout.log"
$stderrPath = Join-Path $workingDirectory "demo-worker.stderr.log"
Remove-Item -LiteralPath $stdoutPath, $stderrPath, $resolvedProcessInfo, $dialogLog `
    -Force -ErrorAction SilentlyContinue
$arguments = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$PSCommandPath`"",
    "-WorkbookPath", "`"$resolvedWorkbook`"",
    "-MacroName", "`"$MacroName`"",
    "-ProcessInfoPath", "`"$resolvedProcessInfo`"",
    "-MaxOrderingBenchmarkSeconds", $MaxOrderingBenchmarkSeconds,
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
        throw "Demo execution exceeded the hard $TimeoutSeconds-second deadline."
    }
    $process.WaitForExit()
    if ((Test-Path -LiteralPath $stderrPath) -and
        (Get-Item -LiteralPath $stderrPath).Length -gt 0) {
        Get-Content -LiteralPath $stderrPath | ForEach-Object {
            [Console]::Error.WriteLine($_)
        }
        throw "Demo execution worker failed."
    }
    Get-Content -LiteralPath $stdoutPath | Write-Output
}
finally {
    if (-not $process.HasExited) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    $process.Dispose()
}
