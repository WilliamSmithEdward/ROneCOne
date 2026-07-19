Attribute VB_Name = "EventsDemoUsage"
Option Explicit

' This executable tutorial demonstrates typed events over multicast Actions.

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
    Dim changed As ROneCOne
    Dim firstHandler As ROneCOne
    Dim removed As Boolean
    Dim secondHandler As ROneCOne

    Set firstHandler = ROneCOne.Action( _
        "EventsDemoUsage.DemoRecordFirst").Takes(vbString)
    Set secondHandler = ROneCOne.Action( _
        "EventsDemoUsage.DemoRecordSecond").Takes(vbString)
    Set changed = ROneCOne.EventOf(vbString) _
        .Subscribe(firstHandler) _
        .Subscribe(secondHandler)

    mTrace = vbNullString
    changed.Emit "ready"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = mTrace
        .Range("E7").Value2 = changed.HandlerCount
    End With

    removed = changed.Unsubscribe(secondHandler)
    mTrace = vbNullString
    changed.Emit "again"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E8").Value2 = removed
        .Range("E9").Value2 = mTrace
    End With
End Sub

Public Sub DemoRecordFirst(ByVal value As Variant)
    mTrace = mTrace & "first:" & CStr(value) & "|"
End Sub

Public Sub DemoRecordSecond(ByVal value As Variant)
    mTrace = mTrace & "second:" & CStr(value) & "|"
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
