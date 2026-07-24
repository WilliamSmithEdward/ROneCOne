Attribute VB_Name = "ProcessDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: awaitable command lines
' ----------------------------------------------------------------------------
' This demo runs a handful of ordinary cmd.exe built-ins and reads their
' results. Nothing is installed and nothing goes online; the commands are
' echo, exit, and cd. The point is what VBA's own Shell function never gave
' you: an exit code, both output streams, and an awaitable handle, with
' Excel staying responsive while the command runs beside it.
'
' The shape mirrors System.Diagnostics.Process: a command that runs and
' fails does not raise; it reports through ExitCode and StandardError. Only
' a command that cannot start at all raises a typed error. Several commands
' can run at the same time and one WhenAll collects every result.
'
' To run it: press Alt+F8, choose RunROneCOneProcessDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_COMMANDS As Long = 3
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mTrace As String

Public Sub RunROneCOneProcessDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteProcessExamples
    RunProcessBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteProcessExamples()
    Dim emptyError As Long
    Dim hello As ROneCOne
    Dim located As ROneCOne
    Dim results As ROneCOne
    Dim warning As ROneCOne

    ' Step 1: run one command and await it. The result object carries the
    ' exit code and everything the command printed, already captured.
    Set hello = ROneCOne.Process.RunAsync("echo hello from ROneCOne").Await

    ' Step 2: standard error travels separately from standard output, so a
    ' warning does not pollute the data you actually wanted.
    Set warning = ROneCOne.Process.RunAsync("echo be careful 1>&2").Await

    ' Step 3: a working directory applies to just that command. Excel's own
    ' current folder never changes.
    Set located = ROneCOne.Process.RunAsync("cd", ThisWorkbook.Path).Await

    ' Step 4: commands overlap. Both processes run at the same time outside
    ' Excel while one WhenAll collects the results in order.
    Set results = ROneCOne.Task.WhenAll( _
        ROneCOne.Process.RunAsync("echo first"), _
        ROneCOne.Process.RunAsync("echo second")).Await

    ' Step 5: failure stays honest. An empty command is a contract error you
    ' can trap by its published number; the trap below records its verdict.
    mTrace = "unexpected acceptance"
    On Error Resume Next
    ROneCOne.Process.RunAsync "   "
    emptyError = Err.Number
    On Error GoTo 0
    If emptyError = ROneCOne.InvalidArgumentError Then
        mTrace = "empty command rejected"
    End If

    ' Each line reads one result and writes it to the Examples sheet, so
    ' every behavior above shows its answer next to what the sheet expects.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = (hello.ExitCode = 0)
        .Range("E7").Value2 = _
            (InStr(1, hello.StandardOutput, "hello from ROneCOne") > 0)
        .Range("E8").Value2 = _
            ROneCOne.Process.RunAsync("exit 7").Await.ExitCode
        .Range("E9").Value2 = _
            (InStr(1, warning.StandardError, "be careful") > 0)
        .Range("E10").Value2 = (InStr(1, located.StandardOutput, _
            ThisWorkbook.Path, vbTextCompare) > 0)
        .Range("E11").Value2 = results.Count
        .Range("E12").Value2 = (ROneCOne.Process.RunAsync( _
            "definitely_not_a_command_xyz").Await.ExitCode <> 0)
        .Range("E13").Value2 = mTrace
    End With
End Sub

Private Sub RunProcessBenchmark()
    Dim elapsed As Double
    Dim results As ROneCOne
    Dim started As Double

    ' Three commands start together and one WhenAll collects them: the wall
    ' time is close to the slowest single command, not the sum, because the
    ' processes run beside Excel while the scheduler polls.
    started = Timer
    Set results = ROneCOne.Task.WhenAll( _
        ROneCOne.Process.RunAsync("echo one"), _
        ROneCOne.Process.RunAsync("echo two"), _
        ROneCOne.Process.RunAsync("echo three")).Await
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_COMMANDS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = results.Count
    End With
End Sub

Private Sub MarkDemoPassed()
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "PASS"
        .Range("B14").ClearContents
    End With
End Sub

Private Sub MarkDemoFailed(ByVal errorNumber As Long, ByVal description As String)
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "ERROR"
        .Range("B14").Value2 = CStr(errorNumber) & ": " & description
    End With
End Sub

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
