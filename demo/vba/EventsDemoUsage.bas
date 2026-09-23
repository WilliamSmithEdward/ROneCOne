Attribute VB_Name = "EventsDemoUsage"
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
' ROneCOne tutorial: typed events
' ----------------------------------------------------------------------------
' An "event" is an announcement one part of your workbook makes so that other
' parts can react. You subscribe interested handlers to the event; then, each
' time you "emit" it, every subscriber runs with the message you sent.
'
' Here an order-status change announces itself, and two handlers react: one
' updates a dashboard, the other writes an audit line. The event is typed: it
' only accepts handlers with the right argument, so a mismatched handler is
' rejected up front rather than failing mid-run.
'
' To run it: press Alt+F8, choose RunROneCOneEventsDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ITERATIONS As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mCount As Long
Private mTrace As String

Public Sub RunROneCOneEventsDemo()
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    WriteEventExamples
    RunEventBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteEventExamples()
    Dim orderStatusChanged As ROneCOne
    Dim removed As Boolean
    Dim dashboard As ROneCOne
    Dim audit As ROneCOne

    ' Wrap two ordinary procedures as handlers. Takes(vbString) says each one
    ' expects a single text message, so only compatible handlers can subscribe.
    Set dashboard = ROneCOne.Action( _
        "EventsDemoUsage.UpdateDashboard").Takes(vbString)
    Set audit = ROneCOne.Action( _
        "EventsDemoUsage.WriteAudit").Takes(vbString)

    ' Create the event and subscribe both handlers. From now on, one Emit call
    ' delivers the message to every subscriber in the order they were added.
    Set orderStatusChanged = ROneCOne.EventOf(vbString) _
        .Subscribe(dashboard) _
        .Subscribe(audit)

    mTrace = vbNullString
    orderStatusChanged.Emit "Order 1042 shipped"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = mTrace
        .Range("E7").Value2 = orderStatusChanged.HandlerCount
    End With

    ' Unsubscribe removes one handler and reports whether it was found. The next
    ' Emit reaches only the dashboard; auditing has been switched off cleanly.
    removed = orderStatusChanged.Unsubscribe(audit)
    mTrace = vbNullString
    orderStatusChanged.Emit "Order 1043 delayed"
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E8").Value2 = removed
        .Range("E9").Value2 = mTrace
    End With
End Sub

Public Sub UpdateDashboard(ByVal msg As Variant)
    mTrace = "Dashboard updated"
End Sub

Public Sub WriteAudit(ByVal msg As Variant)
    mTrace = mTrace & "; audit written"
End Sub

Public Sub DemoCountEvent(ByVal itemValue As Variant)
    mCount = mCount + CLng(itemValue)
End Sub

Private Sub RunEventBenchmark()
    Dim changed As ROneCOne
    Dim handler As ROneCOne
    Dim idx As Long
    Dim started As Double

    Set handler = ROneCOne.Action( _
        "EventsDemoUsage.DemoCountEvent").Takes(vbLong)
    Set changed = ROneCOne.EventOf(vbLong).Subscribe(handler)
    mCount = 0
    started = Timer
    For idx = 1 To BENCHMARK_ITERATIONS
        changed.Emit 1
    Next idx

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
