Attribute VB_Name = "EventsDemoUsage"
Option Explicit

' This tutorial shows one order update reaching several workbook features.
' The event checks every handler before anything is subscribed or run.

Private Const BENCHMARK_ITERATIONS As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mCount As Long
Private mTrace As String

Public Sub RunROneCOneEventsDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteEventExamples
    RunEventBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteEventExamples()
    Dim orderStatusChanged As ROneCOne
    Dim removed As Boolean
    Dim updateDashboard As ROneCOne
    Dim writeAudit As ROneCOne

    ' Turn two ordinary procedures into checked String event handlers.
    Set updateDashboard = ROneCOne.Action( _
        "EventsDemoUsage.UpdateDashboard").Takes(vbString)
    Set writeAudit = ROneCOne.Action( _
        "EventsDemoUsage.WriteAudit").Takes(vbString)

    ' Every Emit now updates the dashboard and records the same message.
    Set orderStatusChanged = ROneCOne.EventOf(vbString) _
        .Subscribe(updateDashboard) _
        .Subscribe(writeAudit)

    mTrace = vbNullString
    orderStatusChanged.Emit "Order 1042 shipped"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = mTrace
        .Range("E7").Value2 = orderStatusChanged.HandlerCount
    End With

    ' Stop auditing future messages without changing the dashboard handler.
    removed = orderStatusChanged.Unsubscribe(writeAudit)
    mTrace = vbNullString
    orderStatusChanged.Emit "Order 1043 delayed"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E8").Value2 = removed
        .Range("E9").Value2 = mTrace
    End With
End Sub

Public Sub UpdateDashboard(ByVal message As Variant)
    mTrace = "Dashboard updated"
End Sub

Public Sub WriteAudit(ByVal message As Variant)
    mTrace = mTrace & "; audit written"
End Sub

Public Sub DemoCountEvent(ByVal value As Variant)
    mCount = mCount + CLng(value)
End Sub

Private Sub RunEventBenchmark()
    Dim changed As ROneCOne
    Dim handler As ROneCOne
    Dim index As Long
    Dim started As Double

    Set handler = ROneCOne.Action( _
        "EventsDemoUsage.DemoCountEvent").Takes(vbLong)
    Set changed = ROneCOne.EventOf(vbLong).Subscribe(handler)
    mCount = 0
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        changed.Emit CLng(1)
    Next index

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = mCount
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
