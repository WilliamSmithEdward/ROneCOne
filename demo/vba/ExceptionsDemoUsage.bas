Attribute VB_Name = "ExceptionsDemoUsage"
Option Explicit

' This executable tutorial demonstrates structured exception flow over Actions.

Private Const BENCHMARK_ITERATIONS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mTrace As String

Public Sub RunROneCOneExceptionsDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteExceptionExamples
    RunExceptionBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteExceptionExamples()
    Dim attempt As ROneCOne
    Dim cleanup As ROneCOne
    Dim errorHandler As ROneCOne
    Dim errorNumber As Long
    Dim failingWork As ROneCOne
    Dim successfulWork As ROneCOne

    Set failingWork = ROneCOne.Action( _
        ROneCOne.Value(1).Divide(CLng(0))).Takes()
    Set errorHandler = ROneCOne.Action( _
        "ExceptionsDemoUsage.DemoCatch").Takes(ROneCOne.Exception)
    Set cleanup = ROneCOne.Action( _
        "ExceptionsDemoUsage.DemoFinally").Takes()

    mTrace = vbNullString
    Set attempt = ROneCOne.Try(failingWork) _
        .Catch(errorHandler) _
        .Finally(cleanup)
    attempt.Execute
    ThisWorkbook.Worksheets(EXAMPLES_SHEET).Range("E6").Value2 = mTrace

    mTrace = vbNullString
    Set attempt = ROneCOne.Try(failingWork) _
        .Catch(vbObjectError + 1, errorHandler) _
        .Finally(cleanup)
    On Error Resume Next
    attempt.Execute
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E7").Value2 = errorNumber
        .Range("E8").Value2 = mTrace
    End With

    Set successfulWork = ROneCOne.Action( _
        "ExceptionsDemoUsage.DemoWork").Takes()
    Set attempt = ROneCOne.Try(successfulWork).Finally(cleanup)
    mTrace = vbNullString
    attempt.Execute
    ThisWorkbook.Worksheets(EXAMPLES_SHEET).Range("E9").Value2 = mTrace
End Sub

Public Sub DemoCatch(ByVal errorInfo As Variant)
    mTrace = mTrace & "caught:" & CStr(errorInfo.ErrorNumber) & "|"
End Sub

Public Sub DemoFinally()
    mTrace = mTrace & "finally|"
End Sub

Public Sub DemoWork()
    mTrace = mTrace & "work|"
End Sub

Private Sub RunExceptionBenchmark()
    Dim attempt As ROneCOne
    Dim index As Long
    Dim started As Double
    Dim work As ROneCOne

    Set work = ROneCOne.Action("ExceptionsDemoUsage.DemoWork").Takes()
    Set attempt = ROneCOne.Try(work)
    mTrace = vbNullString
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        attempt.Execute
    Next index

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = Len(mTrace)
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
