param(
    [string]$WorkbookPath = "tests\output\ROneCOne_DelegateTests.xlsm",
    [ValidateRange(5, 300)]
    [int]$TimeoutSeconds = 30,
    [ValidateRange(0.01, 60)]
    [double]$MaxBenchmarkSeconds = 0.5,
    [ValidateRange(0.01, 60)]
    [double]$MaxCollectionBenchmarkSeconds = 0.75
)

$ErrorActionPreference = "Stop"

# Start-Process rejects environments that contain both Path and PATH. Some
# automation hosts expose both spellings even on Windows, so collapse them in
# this task process before launching the bounded worker.
$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$outputDirectory = Join-Path $PSScriptRoot "..\tests\output"
$workerOutput = Join-Path $outputDirectory "excel-test-worker.stdout.log"
$workerError = Join-Path $outputDirectory "excel-test-worker.stderr.log"
$processInfo = Join-Path $outputDirectory "excel-test-processes.json"
$dialogLog = Join-Path $outputDirectory "vbe-dialogs.jsonl"
$worker = $null

function Stop-TaskProcess {
    param([Nullable[int]]$ProcessId)

    if ($null -eq $ProcessId -or $ProcessId.Value -le 0) {
        return
    }
    Stop-Process -Id $ProcessId.Value -Force -ErrorAction SilentlyContinue
}

function Write-PopupSummary {
    if (-not (Test-Path -LiteralPath $dialogLog)) {
        return
    }

    [Console]::Error.WriteLine("Captured Office/VBE surfaces:")
    Get-Content -LiteralPath $dialogLog | ForEach-Object {
        try {
            $record = $_ | ConvertFrom-Json
            $text = @($record.child_text | Where-Object { $_ }) -join " | "
            $summary = "  class={0}; title={1}; action={2}; text={3}" -f `
                $record.class_name,
                $record.title,
                $record.dismissal_action,
                $text
            [Console]::Error.WriteLine($summary)
            $selections = @(
                $record.automation |
                    Where-Object { $_.selected_text } |
                    ForEach-Object { $_.selected_text }
            ) -join " | "
            if ($selections.Length -gt 0) {
                [Console]::Error.WriteLine("    selected-code=$selections")
            }
        }
        catch {
            [Console]::Error.WriteLine("  unreadable popup record")
        }
    }
}

[System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
Remove-Item -LiteralPath $workerOutput, $workerError, $processInfo, $dialogLog `
    -Force -ErrorAction SilentlyContinue

try {
    $workerScript = Join-Path $PSScriptRoot "run_excel_tests_worker.ps1"
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$workerScript`"",
        "-WorkbookPath", "`"$resolvedWorkbook`"",
        "-ProcessInfoPath", "`"$processInfo`"",
        "-MaxBenchmarkSeconds", $MaxBenchmarkSeconds,
        "-MaxCollectionBenchmarkSeconds", $MaxCollectionBenchmarkSeconds
    )
    $worker = Start-Process `
        -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList $arguments `
        -WindowStyle Hidden `
        -RedirectStandardOutput $workerOutput `
        -RedirectStandardError $workerError `
        -PassThru

    if (-not $worker.WaitForExit($TimeoutSeconds * 1000)) {
        $ownedProcesses = $null
        if (Test-Path -LiteralPath $processInfo) {
            try {
                $ownedProcesses = Get-Content -Raw -LiteralPath $processInfo | ConvertFrom-Json
            }
            catch {
                $ownedProcesses = $null
            }
        }

        if ($null -ne $ownedProcesses) {
            Stop-TaskProcess -ProcessId $ownedProcesses.excel_process_id
            Stop-TaskProcess -ProcessId $ownedProcesses.watcher_process_id
        }
        Stop-TaskProcess -ProcessId $worker.Id
        Write-PopupSummary
        throw "Live Excel test exceeded the hard $TimeoutSeconds-second deadline; task-owned processes were terminated."
    }

    $worker.WaitForExit()
    $worker.Refresh()
    $workerExitCode = $worker.ExitCode
    if ($null -eq $workerExitCode) {
        $hasWorkerError = (Test-Path -LiteralPath $workerError) -and `
            (Get-Item -LiteralPath $workerError).Length -gt 0
        $workerExitCode = if ($hasWorkerError) { -1 } else { 0 }
    }
    if (Test-Path -LiteralPath $workerOutput) {
        Get-Content -LiteralPath $workerOutput | Write-Output
    }
    if (Test-Path -LiteralPath $workerError) {
        Get-Content -LiteralPath $workerError | ForEach-Object {
            [Console]::Error.WriteLine($_)
        }
    }
    if ($workerExitCode -ne 0) {
        Write-PopupSummary
        throw "Live Excel worker failed with exit code $workerExitCode."
    }
}
finally {
    if ($null -ne $worker) {
        if (-not $worker.HasExited) {
            Stop-TaskProcess -ProcessId $worker.Id
        }
        $worker.Dispose()
    }
}
