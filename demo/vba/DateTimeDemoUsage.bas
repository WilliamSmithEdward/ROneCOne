Attribute VB_Name = "DateTimeDemoUsage"
Option Explicit

' ROneCOne 1.10.1, released 2026-09-23
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
' ROneCOne tutorial: dates, times, and durations
' ----------------------------------------------------------------------------
' This demo never touches the network. It reads the timestamps real APIs
' send: full ISO 8601 with offsets, and raw Unix epoch numbers. A value works
' like DateTimeOffset in C#: one instant, viewed at one clock offset, so two
' views of the same moment compare equal even when their clocks differ.
'
' The surface mirrors what C# programmers know: DateTime.Parse and TryParse,
' UtcNow, FromUnixTimeSeconds and ToUnixTimeSeconds, AddMonths with .NET's
' month-end clamping, Subtract yielding a TimeSpan, and round-trip
' ToIsoString text. Windows performs every local conversion, with daylight
' saving applied per instant; the runtime hard-codes no offsets and no rules.
'
' To run it: press Alt+F8, choose RunROneCOneDateTimeDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneDateTimeDemo()
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    WriteDateTimeExamples
    RunDateTimeBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteDateTimeExamples()
    Dim dueAt As ROneCOne
    Dim missError As Long
    Dim posted As ROneCOne
    Dim startAt As ROneCOne
    Dim trace As String

    ' One parsed instant serves several examples. The clock reads 18:30 at
    ' +02:00; the instant itself is 16:30 universal time.
    Set posted = ROneCOne.DateTime.Parse("2026-07-24T18:30:05.123+02:00")
    Set startAt = ROneCOne.DateTime.Parse("2026-07-24T10:00:00Z")
    Set dueAt = ROneCOne.DateTime.Parse("2026-07-24T12:30:00Z")

    ' Impossible text raises the typed FormatError instead of guessing.
    trace = "unexpected success"
    On Error Resume Next
    ROneCOne.DateTime.Parse "2026-02-30"
    missError = Err.Number
    On Error GoTo 0
    If missError = ROneCOne.FormatError Then trace = "impossible date refused"

    ' Cells that would look like dates, times, or plain numbers carry a word
    ' beside the value, because Excel coerces bare date-shaped text on write.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = posted.Hour
        .Range("E7").Value2 = "utc " & posted.ToUniversalTime.ToIsoString
        .Range("E8").Value2 = "epoch " & _
            ROneCOne.DateTime.FromUnixTimeSeconds(1784910605).ToIsoString
        .Range("E9").Value2 = CStr(posted.ToUnixTimeSeconds) & " seconds"
        .Range("E10").Value2 = (posted.CompareTo(ROneCOne.DateTime.Parse( _
            "2026-07-24T16:30:05.123Z")) = 0)
        .Range("E11").Value2 = (ROneCOne.DateTime.Parse( _
            "2026-01-31T12:00:00Z").AddMonths(1).ToIsoString = _
            "2026-02-28T12:00:00Z")
        .Range("E12").Value2 = dueAt.Subtract(startAt).TotalHours
        .Range("E13").Value2 = "lasts " & _
            ROneCOne.TimeSpan.FromMinutes(90).ToString
        .Range("E14").Value2 = trace
    End With
End Sub

Private Sub RunDateTimeBenchmark()
    Dim elapsed As Double
    Dim idx As Long
    Dim moment As ROneCOne
    Dim roundTrips As Long
    Dim stamp As String
    Dim started As Double

    ' Parse a thousand ISO timestamps and format each one back, counting the
    ' exact round trips. Everything is integer arithmetic on milliseconds.
    started = Timer
    For idx = 1 To BENCHMARK_ROWS
        stamp = "2026-07-24T" & Right$("0" & CStr(idx Mod 24), 2) & _
            ":30:05." & Right$("00" & CStr((idx Mod 999) + 1), 3) & "Z"
        Set moment = ROneCOne.DateTime.Parse(stamp)
        If moment.ToIsoString = stamp Then roundTrips = roundTrips + 1
    Next idx
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = roundTrips
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
