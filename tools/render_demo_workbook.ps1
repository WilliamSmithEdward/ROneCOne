param(
    [string]$WorkbookPath = "demo\ROneCOne_Delegates_Demo.xlsm",
    [string]$OutputPrefix = "delegates",
    [string]$NodeExecutable = "",
    [string]$NodeModulesPath = "",
    [ValidateRange(10, 180)]
    [int]$TimeoutSeconds = 45,
    [ValidateRange(0, 15)]
    [int]$LateExcelGraceSeconds = 8
)

$ErrorActionPreference = "Stop"

$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

if ([string]::IsNullOrWhiteSpace($NodeExecutable)) {
    $nodePath = (Get-Command node -ErrorAction Stop).Source
} else {
    $nodePath = (Resolve-Path -LiteralPath $NodeExecutable).Path
}

if (-not [string]::IsNullOrWhiteSpace($NodeModulesPath)) {
    $nodeModules = (Resolve-Path -LiteralPath $NodeModulesPath).Path
} elseif (-not [string]::IsNullOrWhiteSpace($env:NODE_PATH)) {
    $nodeModules = $env:NODE_PATH
} else {
    $localNodeModules = Join-Path $PSScriptRoot "..\node_modules"
    if (-not (Test-Path -LiteralPath $localNodeModules)) {
        throw "Artifact-tool modules were not found. Set NODE_PATH or pass -NodeModulesPath."
    }
    $nodeModules = (Resolve-Path -LiteralPath $localNodeModules).Path
}

$scriptPath = Join-Path $PSScriptRoot "render_demo_workbook.cjs"
$workingDirectory = Join-Path $PSScriptRoot "..\demo\.working"
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$workbookLock = Join-Path `
    (Split-Path -Parent $resolvedWorkbook) `
    ("~`$" + (Split-Path -Leaf $resolvedWorkbook))
$stdoutPath = Join-Path $workingDirectory "render-worker.stdout.log"
$stderrPath = Join-Path $workingDirectory "render-worker.stderr.log"
$beforeExcelIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$trackedExcelIds = [System.Collections.Generic.HashSet[int]]::new()
$watchers = @{}
$nodeProcess = $null
$timedOut = $false

function Start-PopupWatcher {
    param([int]$ExcelProcessId)

    if ($watchers.ContainsKey($ExcelProcessId)) {
        return
    }
    $watcherScript = Join-Path $PSScriptRoot "watch_vbe_dialogs.ps1"
    $dialogLog = Join-Path $workingDirectory "render-dialogs-$ExcelProcessId.jsonl"
    $stopPath = Join-Path $workingDirectory "render-watcher-$ExcelProcessId.stop"
    Remove-Item -LiteralPath $dialogLog, $stopPath -Force -ErrorAction SilentlyContinue
    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$watcherScript`"",
        "-ExcelProcessId", $ExcelProcessId,
        "-LogPath", "`"$dialogLog`"",
        "-StopPath", "`"$stopPath`"",
        "-TimeoutSeconds", 120,
        "-DismissKnownDialogs",
        "-TerminateOnBreakMode"
    )
    $watcher = Start-Process `
        -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList $arguments `
        -WindowStyle Hidden `
        -PassThru
    $watchers[$ExcelProcessId] = [pscustomobject]@{
        Process = $watcher
        LogPath = $dialogLog
        StopPath = $stopPath
    }
}

function Find-RendererExcelProcesses {
    if ($null -eq $nodeProcess) {
        return @()
    }

    $newExcelProcesses = @(
        Get-Process EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $beforeExcelIds -notcontains $_.Id }
    )
    if ($newExcelProcesses.Count -eq 0) {
        return @()
    }

    $parentMatchedIds = @(
        Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue |
            Where-Object {
                $beforeExcelIds -notcontains [int]$_.ProcessId -and
                [int]$_.ParentProcessId -eq $nodeProcess.Id
            } |
            ForEach-Object { [int]$_.ProcessId }
    )
    if ($parentMatchedIds.Count -gt 0) {
        return @(
            $parentMatchedIds | ForEach-Object {
                [pscustomobject]@{
                    ProcessId = $_
                    OwnershipBasis = "renderer child process"
                }
            }
        )
    }

    if (-not (Test-Path -LiteralPath $workbookLock)) {
        return @()
    }

    $lockCreated = (Get-Item -LiteralPath $workbookLock).CreationTimeUtc
    $lockMatchedProcesses = @(
        $newExcelProcesses | Where-Object {
            try {
                [Math]::Abs(
                    ($_.StartTime.ToUniversalTime() - $lockCreated).TotalSeconds
                ) -le 3
            }
            catch {
                $false
            }
        }
    )
    if ($lockMatchedProcesses.Count -ne 1) {
        return @()
    }

    return @(
        [pscustomobject]@{
            ProcessId = [int]$lockMatchedProcesses[0].Id
            OwnershipBasis = "lock-matched Excel process"
        }
    )
}

[System.IO.Directory]::CreateDirectory($workingDirectory) | Out-Null
Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
$previousNodePath = $env:NODE_PATH
$env:NODE_PATH = $nodeModules

try {
    $nodeProcess = Start-Process `
        -FilePath $nodePath `
        -ArgumentList "`"$scriptPath`" `"$resolvedWorkbook`" `"$OutputPrefix`"" `
        -WorkingDirectory (Join-Path $PSScriptRoot "..") `
        -WindowStyle Hidden `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -PassThru

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $nodeExitedAt = $null
    while ([DateTime]::UtcNow -lt $deadline) {
        foreach ($candidate in @(Find-RendererExcelProcesses)) {
            $excelId = [int]$candidate.ProcessId
            if ($trackedExcelIds.Add($excelId)) {
                Start-PopupWatcher -ExcelProcessId $excelId
            }
        }

        if ($nodeProcess.HasExited) {
            if ($null -eq $nodeExitedAt) {
                $nodeExitedAt = [DateTime]::UtcNow
            }
            if ([DateTime]::UtcNow -ge $nodeExitedAt.AddSeconds($LateExcelGraceSeconds)) {
                break
            }
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $nodeProcess.HasExited) {
        $timedOut = $true
    }
}
finally {
    foreach ($entry in $watchers.Values) {
        Set-Content -LiteralPath $entry.StopPath -Value "stop"
        if (-not $entry.Process.WaitForExit(2000)) {
            Stop-Process -Id $entry.Process.Id -Force -ErrorAction SilentlyContinue
        }
        $entry.Process.Dispose()
    }
    foreach ($excelId in $trackedExcelIds) {
        Stop-Process -Id $excelId -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $nodeProcess) {
        if (-not $nodeProcess.HasExited) {
            Stop-Process -Id $nodeProcess.Id -Force -ErrorAction SilentlyContinue
        }
        $nodeProcess.Dispose()
    }
    Start-Sleep -Milliseconds 250
    Remove-Item -LiteralPath $workbookLock -Force -ErrorAction SilentlyContinue
    $env:NODE_PATH = $previousNodePath
}

if ($timedOut) {
    throw "Demo rendering exceeded the hard $TimeoutSeconds-second deadline."
}
if ((Test-Path -LiteralPath $stderrPath) -and
    (Get-Item -LiteralPath $stderrPath).Length -gt 0) {
    Get-Content -LiteralPath $stderrPath | ForEach-Object {
        [Console]::Error.WriteLine($_)
    }
    throw "Demo rendering worker failed."
}
foreach ($entry in $watchers.Values) {
    if (Test-Path -LiteralPath $entry.LogPath) {
        $modalRecords = @(
            Get-Content -LiteralPath $entry.LogPath |
                ConvertFrom-Json |
                Where-Object {
                    $_.class_name -eq "#32770" -or $_.dismissal_action -ne "none"
                }
        )
        if ($modalRecords.Count -gt 0) {
            throw "Office or VBE popup was observed while rendering the demo."
        }
    }
}
Get-Content -LiteralPath $stdoutPath | Write-Output
Write-Output "Renderer-owned Excel processes cleaned: $($trackedExcelIds.Count)"
