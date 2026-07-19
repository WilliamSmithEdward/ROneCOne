Attribute VB_Name = "TasksDemoUsage"
Option Explicit

' This executable tutorial demonstrates cooperative Task workflows in one Excel process.

Private Const BENCHMARK_ITERATIONS As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mProgressTotal As Long
Private mTrace As String

Public Sub RunROneCOneTasksDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteTaskExamples
    RunTaskBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteTaskExamples()
    Dim bounded As ROneCOne
    Dim completion As ROneCOne
    Dim continuation As ROneCOne
    Dim delayed As ROneCOne
    Dim firstTask As ROneCOne
    Dim progress As ROneCOne
    Dim results As ROneCOne
    Dim secondTask As ROneCOne
    Dim source As ROneCOne
    Dim ignored As Variant
    Dim registration As ROneCOne

    Set firstTask = ROneCOne.Task.FromResult(10&)
    Set secondTask = ROneCOne.Task.FromResult(20&)
    Set results = ROneCOne.Task.WhenAll(firstTask, secondTask).Await

    Set delayed = ROneCOne.Task.Delay(5&)
    Set continuation = ROneCOne.Func( _
        "TasksDemoUsage.TaskPlusOne") _
        .Takes(ROneCOne.Task) _
        .Returns(vbLong)

    Set source = ROneCOne.CancellationTokenSource
    mTrace = vbNullString
    Set registration = source.Token.Register(ROneCOne.Action( _
        "TasksDemoUsage.RecordCancellation").Takes)
    source.Cancel
    registration.Dispose

    mProgressTotal = 0
    Set progress = ROneCOne.ProgressOf( _
        vbLong, ROneCOne.Action( _
            "TasksDemoUsage.RecordProgress").Takes(vbLong))
    progress.Report 7&

    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    completion.SetResult 99&

    Set bounded = ROneCOne.Task.Delay(5&).WaitAsync(100&)
    ignored = bounded.Await
    ignored = ROneCOne.Task.YieldOnce.Await

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = ROneCOne.Task.FromResult(42&).Await
        ignored = delayed.Await
        .Range("E7").Value2 = delayed.IsCompleted
        .Range("E8").Value2 = results.JoinText(",")
        .Range("E9").Value2 = firstTask.ContinueWith(continuation).Await
        .Range("E10").Value2 = source.Token.IsCancellationRequested
        .Range("E11").Value2 = mProgressTotal
        .Range("E12").Value2 = completion.Task.Await
        .Range("E13").Value2 = bounded.IsCompleted
        .Range("E14").Value2 = True
    End With
End Sub

Public Function TaskPlusOne(ByVal antecedent As Variant) As Variant
    TaskPlusOne = CLng(antecedent.Result) + 1&
End Function

Public Sub RecordCancellation()
    mTrace = mTrace & "canceled|"
End Sub

Public Sub RecordProgress(ByVal value As Variant)
    mProgressTotal = mProgressTotal + CLng(value)
End Sub

Private Sub RunTaskBenchmark()
    Dim index As Long
    Dim result As Long
    Dim started As Double

    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        result = ROneCOne.Task.FromResult(index).Await
    Next index
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = result
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
