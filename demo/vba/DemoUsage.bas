Attribute VB_Name = "DemoUsage"
Option Explicit

' This module is an executable tutorial for the delegate feature slice.
' The public entry point coordinates focused, independently readable sections.

Private Const BENCHMARK_ITERATIONS As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

' -----------------------------------------------------------------------------
' Entry point
' -----------------------------------------------------------------------------

Public Sub RunROneCOneDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure

    WriteDelegateExamples
    RunDelegateBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

' -----------------------------------------------------------------------------
' Delegate examples
' -----------------------------------------------------------------------------

Private Sub WriteDelegateExamples()
    Dim addValues As ROneCOne
    Dim between As ROneCOne
    Dim doubleValue As ROneCOne
    Dim maximum As ROneCOne
    Dim pipeline As ROneCOne
    Dim safeFalse As ROneCOne
    Dim square As ROneCOne
    Dim worksheetFunctions As Object
    Dim x As ROneCOne
    Dim y As ROneCOne

    ' Build reusable delegates without strings or runtime code generation.
    Set x = ROneCOne.Parameter(vbLong)
    Set y = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)
    Set addValues = ROneCOne.Lambda(x.Add(y), x, y)
    Set between = ROneCOne.Lambda( _
        x.GreaterThan(CLng(10)).AndAlso(x.LessThan(CLng(20))), x)
    Set safeFalse = ROneCOne.Lambda( _
        ROneCOne.Value(False).AndAlso(ROneCOne.Value(1).Divide(0)))

    ' FromMethod adapts a normal object method into the same delegate contract.
    Set worksheetFunctions = Application.WorksheetFunction
    Set maximum = ROneCOne.FromMethod(worksheetFunctions, "Max", 2)
    Set doubleValue = ROneCOne.Lambda(x.Add(x), x)
    Set pipeline = square.PipeTo(doubleValue)

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = square(CLng(9))
        .Range("E7").Value2 = addValues(CLng(6), CLng(7))
        .Range("E8").Value2 = between(CLng(15))
        .Range("E9").Value2 = safeFalse.Run()
        .Range("E10").Value2 = maximum(CLng(4), CLng(7))
        .Range("E11").Value2 = pipeline(CLng(3))
    End With
End Sub

' -----------------------------------------------------------------------------
' Benchmark and reporting
' -----------------------------------------------------------------------------

Private Sub RunDelegateBenchmark()
    Dim index As Long
    Dim lastResult As Variant
    Dim square As ROneCOne
    Dim started As Double
    Dim x As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        lastResult = square(CLng(index))
    Next index

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = lastResult
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
