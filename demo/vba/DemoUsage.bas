Attribute VB_Name = "DemoUsage"
Option Explicit

' This executable tutorial leads with the least-ceremony delegate surface.
' The deeper examples expose the universal invocation kernel and its metadata.

#If Win64 Then
Private Declare PtrSafe Sub CopyPointer Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef destination As LongPtr, _
    ByRef source As LongPtr, _
    ByVal byteCount As LongPtr _
)
#End If

Private Const BENCHMARK_ITERATIONS As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mTrace As String

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
' Frictionless delegate examples
' -----------------------------------------------------------------------------

Private Sub WriteDelegateExamples()
    Dim addValues As ROneCOne
    Dim between As ROneCOne
    Dim combined As ROneCOne
    Dim doubleValue As ROneCOne
    Dim firstAction As ROneCOne
    Dim increment As ROneCOne
    Dim maximum As ROneCOne
    Dim pipeline As ROneCOne
    Dim safeFalse As ROneCOne
    Dim secondAction As ROneCOne
    Dim square As ROneCOne
    Dim value As Long
    Dim workbookAdd As ROneCOne
    Dim worksheetFunctions As Object
    Dim x As ROneCOne
    Dim y As ROneCOne

    ' Var creates typed arguments; AsFunc infers them from the expression tree.
    Set x = ROneCOne.Var(vbLong)
    Set y = ROneCOne.Var(vbLong)
    Set square = x.Multiply(x).AsFunc
    Set addValues = x.Add(y).AsFunc
    Set between = x.AtLeast(CLng(10)) _
        .AndAlso(x.LessThan(CLng(20))) _
        .AsFunc
    Set safeFalse = ROneCOne.Value(False) _
        .AndAlso(ROneCOne.Value(1).Divide(0)) _
        .AsFunc

    ' Func adapts an object method or a standard-module procedure.
    Set worksheetFunctions = Application.WorksheetFunction
    Set maximum = ROneCOne.Func(worksheetFunctions, "Max") _
        .Takes(vbLong, vbLong) _
        .Returns(vbDouble)
    Set workbookAdd = ROneCOne.Func("DemoUsage.DemoAddValues") _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)

    ' Combine creates an immutable multicast Action in invocation order.
    Set firstAction = ROneCOne.Action("DemoUsage.DemoRecordFirst") _
        .Takes(vbString)
    Set secondAction = ROneCOne.Action("DemoUsage.DemoRecordSecond") _
        .Takes(vbString)
    Set combined = ROneCOne.Combine(firstAction, secondAction)
    mTrace = vbNullString
    combined.Execute "value"

    ' NativeAction supplies true ByRef semantics when identity must be preserved.
    value = 41
#If Win64 Then
    Set increment = ROneCOne.NativeAction(DemoNativeIncrementLongAddress) _
        .Takes(ROneCOne.RefOf(vbLong))
    increment.Execute ROneCOne.RefLong(value)
#End If

    Set doubleValue = x.Add(x).AsFunc
    Set pipeline = square.PipeTo(doubleValue)

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = square(CLng(9))
        .Range("E7").Value2 = addValues(CLng(6), CLng(7))
        .Range("E8").Value2 = between(CLng(15))
        .Range("E9").Value2 = safeFalse.Run()
        .Range("E10").Value2 = maximum(CLng(4), CLng(7))
        .Range("E11").Value2 = workbookAdd(CLng(6), CLng(7))
        .Range("E12").Value2 = workbookAdd.DynamicInvoke(Array(CLng(20), CLng(22)))
        .Range("E13").Value2 = mTrace
        .Range("E14").Value2 = value
        .Range("E15").Value2 = pipeline(CLng(3))
        .Range("E16").Value2 = workbookAdd.Signature
    End With
End Sub

' -----------------------------------------------------------------------------
' Demo call targets
' -----------------------------------------------------------------------------

Public Function DemoAddValues( _
    ByVal leftValue As Variant, _
    ByVal rightValue As Variant _
) As Variant
    DemoAddValues = leftValue + rightValue
End Function

Public Sub DemoRecordFirst(ByVal value As Variant)
    mTrace = mTrace & "first:" & CStr(value) & "|"
End Sub

Public Sub DemoRecordSecond(ByVal value As Variant)
    mTrace = mTrace & "second:" & CStr(value) & "|"
End Sub

#If Win64 Then
Public Sub DemoNativeIncrementLong(ByRef value As Long)
    value = value + 1
End Sub

Public Function DemoNativeIncrementLongAddress() As LongPtr
    Dim procedureAddress As LongPtr

    CopyPointer procedureAddress, AddressOf DemoNativeIncrementLong, _
        LenB(procedureAddress)
    DemoNativeIncrementLongAddress = procedureAddress
End Function
#End If

' -----------------------------------------------------------------------------
' Benchmark and reporting
' -----------------------------------------------------------------------------

Private Sub RunDelegateBenchmark()
    Dim index As Long
    Dim lastResult As Variant
    Dim square As ROneCOne
    Dim started As Double
    Dim x As ROneCOne

    Set x = ROneCOne.Var(vbLong)
    Set square = x.Multiply(x).AsFunc
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
