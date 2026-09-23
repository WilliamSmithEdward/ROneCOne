Attribute VB_Name = "ProcessDemoUsage"
Option Explicit

' ROneCOne 1.9.1, released 2026-09-22
'
' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

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
    Dim errNumber As Long

    On Error GoTo DemoFailure
    WriteProcessExamples
    RunProcessBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteProcessExamples()
    Dim emptyError As Long
    Dim hello As ROneCOne
    Dim located As ROneCOne
    Dim echoes As ROneCOne
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
    Set echoes = ROneCOne.Task.WhenAll( _
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
        .Range("E11").Value2 = echoes.Count
        .Range("E12").Value2 = (ROneCOne.Process.RunAsync( _
            "definitely_not_a_command_xyz").Await.ExitCode <> 0)
        .Range("E13").Value2 = mTrace
        ' Step 6: feed a command standard input. The lines go in unsorted
        ' and sort hands them back alphabetized, apple before banana.
        .Range("E14").Value2 = (InStr(1, ROneCOne.Process.RunAsync("sort", _
            , , "banana" & vbCrLf & "apple" & vbCrLf).Await.StandardOutput, _
            "apple") < InStr(1, ROneCOne.Process.RunAsync("sort", , , _
            "banana" & vbCrLf & "apple" & vbCrLf).Await.StandardOutput, _
            "banana"))
        ' Step 7: a session is a conversation. One process stays alive and
        ' answers twice, which RunAsync cannot do because it waits for the
        ' command to finish before you see anything.
        .Range("E15").Value2 = SessionConversation()
        ' Step 8: sort reads until end of input, so CloseInput is what lets
        ' it finish and report its exit code.
        .Range("E16").Value2 = SessionSortedFirstLine()
    End With
End Sub

Private Function SessionConversation() As String
    Dim alphaLine As String
    Dim betaLine As String
    Dim session As ROneCOne

    Set session = ROneCOne.Process.StartSession("echo ready")
    session.WriteLineAsync("echo alpha").Await
    alphaLine = SessionLineWith(session, "alpha")
    session.WriteLineAsync("echo beta").Await
    betaLine = SessionLineWith(session, "beta")
    session.WriteLineAsync("exit 0").Await
    session.WaitForExitAsync.Await
    If Len(alphaLine) > 0 And Len(betaLine) > 0 Then
        SessionConversation = alphaLine & " then " & betaLine
    Else
        SessionConversation = "the session did not answer twice"
    End If
End Function

Private Function SessionSortedFirstLine() As String
    Dim session As ROneCOne

    Set session = ROneCOne.Process.StartSession("sort")
    session.WriteLineAsync("banana").Await
    session.WriteLineAsync("apple").Await
    session.CloseInput
    SessionSortedFirstLine = SessionLineWith(session, "apple")
    session.WaitForExitAsync.Await
End Function

' Reads bounded lines until the wanted word appears. A prompt arrives with
' no trailing newline, so every read carries a deadline and resolves to
' empty text rather than waiting on a line that will never come. cmd.exe
' prints its prompt ahead of the answer on the same line, so the match is
' returned from the fragment rather than as the whole line.
Private Function SessionLineWith( _
    ByVal session As ROneCOne, _
    ByVal fragment As String _
) As String
    Dim attempt As Long
    Dim lineText As String
    Dim pos As Long

    For attempt = 1 To 12
        lineText = session.ReadLineAsync(4000).Await
        pos = InStr(1, lineText, fragment)
        If pos > 0 Then
            SessionLineWith = Trim$(Mid$(lineText, pos))
            Exit Function
        End If
        If Len(lineText) = 0 And session.HasExited Then Exit For
    Next attempt
End Function

Private Sub RunProcessBenchmark()
    Dim elapsed As Double
    Dim runOutcomes As ROneCOne
    Dim started As Double

    ' Three commands start together and one WhenAll collects them: the wall
    ' time is close to the slowest single command, not the sum, because the
    ' processes run beside Excel while the scheduler polls.
    started = Timer
    Set runOutcomes = ROneCOne.Task.WhenAll( _
        ROneCOne.Process.RunAsync("echo one"), _
        ROneCOne.Process.RunAsync("echo two"), _
        ROneCOne.Process.RunAsync("echo three")).Await
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_COMMANDS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = runOutcomes.Count
    End With
End Sub

Private Sub MarkDemoPassed()
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "PASS"
        .Range("B14").ClearContents
    End With
End Sub

Private Sub MarkDemoFailed(ByVal errNumber As Long, ByVal errDescription As String)
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "ERROR"
        .Range("B14").Value2 = CStr(errNumber) & ": " & errDescription
    End With
End Sub

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
