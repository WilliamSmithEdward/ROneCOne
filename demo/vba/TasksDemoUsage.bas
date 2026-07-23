Attribute VB_Name = "TasksDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: tasks you can await
' ----------------------------------------------------------------------------
' A "task" is a piece of work you start and then collect the answer from later,
' the way you might put a kettle on and come back when it boils. You start the
' work with Task.Run, keep going, and call Await when you actually need the
' result. Tasks also let you run several pieces together, stop work you no
' longer need, follow its progress, and put a time limit on waiting.
'
' One important thing to know: these tasks take turns on Excel's own thread
' rather than running on separate processor cores at the same instant. They
' cooperate, pausing to let Excel breathe, which keeps your workbook responsive
' and your data safe. They are about coordination and staying responsive, not
' about raw speed from parallel cores.
'
' This demo runs two forecasts together, turns their numbers into a sentence,
' counts open orders, cancels work on request, tracks progress, and waits with
' a timeout. Read top to bottom; each block names what it produces.
'
' To run it: press Alt+F8, choose RunROneCOneTasksDemo, and click Run.
' ============================================================================

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
    Dim yielded As ROneCOne
    Dim ignored As Variant
    Dim registration As ROneCOne

    ' Describe two small calculations as rules, the same delegates from the
    ' delegates tutorial: a sales forecast (up eight percent) and a reorder
    ' point. Nothing runs yet; these are just the work each task will carry out.
    Set forecastWork = ROneCOne.Value(125000#).Multiply(1.08).AsFunc
    Set reorderWork = ROneCOne.Value(80#).Multiply(1.65).Add(20#).AsFunc

    ' Start both at once. Task.Run kicks off each calculation and hands back a
    ' task standing in for the eventual answer. WhenAll bundles them into one
    ' task that finishes when both are done, keeping the answers in the order
    ' you listed them, so the forecast is always first and the reorder second.
    Set forecastTask = ROneCOne.Task.Run(forecastWork)
    Set reorderTask = ROneCOne.Task.Run(reorderWork)
    Set allWork = ROneCOne.Task.WhenAll(forecastTask, reorderTask)

    ' A continuation is follow-up work that runs once a task finishes. This one
    ' takes the two raw numbers and writes them into a readable sentence, so you
    ' chain "do the work" and "then format the result" without waiting in between.
    Set buildSummary = ROneCOne.Func( _
        "TasksDemoUsage.BuildForecastSummary") _
        .Takes(ROneCOne.Task) _
        .Returns(vbString)
    Set summaryTask = allWork.ContinueWith(buildSummary)
    ' Await is where you finally collect the answer; here it hands back both
    ' numbers together, since allWork bundled the two tasks into one.
    Set results = allWork.Await

    ' A task can also run one of your own procedures. This one reads workbook
    ' data to count open orders, which is exactly why it belongs on Excel's
    ' thread with the rest, taking its turn safely alongside everything else.
    Set countOpenOrders = ROneCOne.Func( _
        "TasksDemoUsage.CountOpenOrders").Takes().Returns(vbLong)
    Set openOrdersTask = ROneCOne.Task.Run(countOpenOrders)

    ' Cancellation is how a button, a timeout, or any other signal politely asks
    ' running work to stop. You register what to do when cancel happens, then
    ' Cancel fires it; here it simply records that the cancellation was seen.
    Set source = ROneCOne.CancellationTokenSource
    mTrace = vbNullString
    Set registration = source.Token.Register(ROneCOne.Action( _
        "TasksDemoUsage.RecordCancellation").Takes)
    source.Cancel
    registration.Dispose

    ' Progress reporting lets long work send updates back as it goes, like a
    ' status bar. This progress channel only accepts whole numbers; each Report
    ' call adds to a running total that a handler keeps for us.
    mProgressTotal = 0
    Set progress = ROneCOne.ProgressOf( _
        vbLong, ROneCOne.Action( _
            "TasksDemoUsage.RecordProgress").Takes(vbLong))
    progress.Report 7

    ' Sometimes the answer arrives from an outside event rather than a running
    ' calculation. A completion source is a task whose result you set by hand
    ' whenever that moment comes; here we simply hand it 99 straight away.
    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    completion.SetResult 99

    ' Waiting, three ways. Delay pauses briefly while letting Excel stay
    ' responsive. WaitAsync wraps a wait in a time limit so it cannot hang
    ' forever. YieldOnce steps aside for a single beat to let Excel catch up.
    Set delayed = ROneCOne.Task.Delay(5)
    Set bounded = ROneCOne.Task.Delay(5).WaitAsync(100)
    ignored = delayed.Await
    ignored = bounded.Await
    Set yielded = ROneCOne.Task.YieldOnce
    ignored = yielded.Await

    ' Collect every answer and write it to the Examples sheet: the two forecast
    ' numbers joined together, the formatted summary, the order count, whether
    ' each wait finished, whether cancellation was requested, the progress total,
    ' and the value handed to the completion source, one result per row.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = results.JoinText(" | ")
        .Range("E7").Value2 = summaryTask.Await
        .Range("E8").Value2 = openOrdersTask.Await
        .Range("E9").Value2 = delayed.IsCompleted
        .Range("E10").Value2 = source.Token.IsCancellationRequested
        .Range("E11").Value2 = mProgressTotal
        .Range("E12").Value2 = completion.Task.Await
        .Range("E13").Value2 = bounded.IsCompleted
        .Range("E14").Value2 = yielded.IsCompleted
    End With
End Sub

Public Function CountOpenOrders() As Variant
    Dim status As Variant

    For Each status In Array( _
        "Open", "Shipped", "Open", "Pending", "Open", "Shipped")
        If status = "Open" Then CountOpenOrders = CLng(CountOpenOrders) + 1
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

    ' Start a task and await its answer, a thousand times over, and time it.
    ' This shows the round trip of scheduling and collecting a result stays
    ' quick enough to use freely, even though tasks take turns on one thread.
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        Set work = ROneCOne.Value(index).Multiply(2).AsFunc
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
