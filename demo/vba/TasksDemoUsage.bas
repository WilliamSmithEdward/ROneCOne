Attribute VB_Name = "TasksDemoUsage"
Option Explicit

' This tutorial runs safe calculations in parallel without opening another Excel.
' Workbook and VBA work uses RunOnExcel so Excel objects never cross threads.

Private Const BENCHMARK_ITERATIONS As Long = 1000
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
    Dim allWork As ROneCOne
    Dim bounded As ROneCOne
    Dim buildSummary As ROneCOne
    Dim completion As ROneCOne
    Dim countOpenOrders As ROneCOne
    Dim delayed As ROneCOne
    Dim forecastTask As ROneCOne
    Dim forecastWork As ROneCOne
    Dim openOrdersTask As ROneCOne
    Dim progress As ROneCOne
    Dim reorderTask As ROneCOne
    Dim reorderWork As ROneCOne
    Dim results As ROneCOne
    Dim source As ROneCOne
    Dim summaryTask As ROneCOne
    Dim ignored As Variant
    Dim registration As ROneCOne

    ' These expression lambdas are pure: they cannot touch Excel, VBA, or COM.
    ' That makes it safe for Task.Run to move them onto Windows worker threads.
    Set forecastWork = ROneCOne.Value(125000#).Multiply(1.08).AsFunc
    Set reorderWork = ROneCOne.Value(80#).Multiply(1.65).Add(20#).AsFunc

    ' Task.Run starts immediately. WhenAll keeps results in the original order.
    Set forecastTask = ROneCOne.Task.Run(forecastWork)
    Set reorderTask = ROneCOne.Task.Run(reorderWork)
    Set allWork = ROneCOne.Task.WhenAll(forecastTask, reorderTask)

    ' A continuation turns the two raw counts into a readable summary.
    Set buildSummary = ROneCOne.Func( _
        "TasksDemoUsage.BuildForecastSummary") _
        .Takes(ROneCOne.Task) _
        .Returns(vbString)
    Set summaryTask = allWork.ContinueWith(buildSummary)
    Set results = allWork.Await

    ' This function reads VBA data, so it deliberately stays on Excel's thread.
    Set countOpenOrders = ROneCOne.Func( _
        "TasksDemoUsage.CountOpenOrders").Takes().Returns(vbLong)
    Set openOrdersTask = ROneCOne.Task.RunOnExcel(countOpenOrders)

    ' Cancellation lets a button or timeout ask cooperative work to stop.
    Set source = ROneCOne.CancellationTokenSource
    mTrace = vbNullString
    Set registration = source.Token.Register(ROneCOne.Action( _
        "TasksDemoUsage.RecordCancellation").Takes)
    source.Cancel
    registration.Dispose

    ' Progress reports are checked to contain Long values.
    mProgressTotal = 0
    Set progress = ROneCOne.ProgressOf( _
        vbLong, ROneCOne.Action( _
            "TasksDemoUsage.RecordProgress").Takes(vbLong))
    progress.Report 7&

    ' A completion source lets an event or callback finish a Task later.
    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    completion.SetResult 99&

    ' Delay yields to Excel, while WaitAsync puts a time limit around waiting.
    Set delayed = ROneCOne.Task.Delay(5&)
    Set bounded = ROneCOne.Task.Delay(5&).WaitAsync(100&)
    ignored = delayed.Await
    ignored = bounded.Await
    ignored = ROneCOne.Task.YieldOnce.Await

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = results.JoinText(" | ")
        .Range("E7").Value2 = summaryTask.Await
        .Range("E8").Value2 = _
            forecastTask.WorkerThreadId <> ROneCOne.CurrentThreadId
        .Range("E9").Value2 = openOrdersTask.Await
        .Range("E10").Value2 = delayed.IsCompleted
        .Range("E11").Value2 = source.Token.IsCancellationRequested
        .Range("E12").Value2 = mProgressTotal
        .Range("E13").Value2 = completion.Task.Await
        .Range("E14").Value2 = bounded.IsCompleted
    End With
End Sub

Public Function CountOpenOrders() As Variant
    Dim status As Variant

    For Each status In Array( _
        "Open", "Shipped", "Open", "Pending", "Open", "Shipped")
        If status = "Open" Then CountOpenOrders = CLng(CountOpenOrders) + 1&
    Next status
End Function

Public Function BuildForecastSummary(ByVal antecedent As Variant) As Variant
    Dim results As ROneCOne

    Set results = antecedent.Result
    BuildForecastSummary = "Forecast " & CStr(results.Item(0)) & _
        "; reorder point " & CStr(results.Item(1))
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
    Dim work As ROneCOne

    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        Set work = ROneCOne.Value(index).Multiply(2&).AsFunc
        result = ROneCOne.Task.Run(work).Await
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
