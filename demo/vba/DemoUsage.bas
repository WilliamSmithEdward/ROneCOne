Attribute VB_Name = "DemoUsage"
Option Explicit

' This tutorial uses delegates to build reusable pricing and notification rules.
' Read the comments from top to bottom; each example starts with the useful result.

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
' Practical delegate examples
' -----------------------------------------------------------------------------

Private Sub WriteDelegateExamples()
    Dim addHandling As ROneCOne
    Dim amount As ROneCOne
    Dim applyDiscount As ROneCOne
    Dim approvalRule As ROneCOne
    Dim calculateTotal As ROneCOne
    Dim increment As ROneCOne
    Dim maximum As ROneCOne
    Dim notify As ROneCOne
    Dim orderNumber As Long
    Dim orderTotal As ROneCOne
    Dim pipeline As ROneCOne
    Dim price As ROneCOne
    Dim safeFalse As ROneCOne
    Dim shipping As ROneCOne
    Dim updateDashboard As ROneCOne
    Dim worksheetFunctions As Object
    Dim writeAudit As ROneCOne

    ' Build pricing rules in place. No extra helper procedure is needed.
    Set amount = ROneCOne.Var(vbLong)
    Set shipping = ROneCOne.Var(vbLong)
    Set price = ROneCOne.Var(vbDouble)
    Set applyDiscount = price.Multiply(0.9).AsFunc
    Set orderTotal = amount.Add(shipping).AsFunc
    Set approvalRule = amount.AtLeast(CLng(100)) _
        .AndAlso(amount.LessThan(CLng(1000))) _
        .AsFunc

    ' AndAlso stops after False, so the unsafe division is never attempted.
    Set safeFalse = ROneCOne.Value(False) _
        .AndAlso(ROneCOne.Value(1).Divide(0)) _
        .AsFunc

    ' Existing Excel methods and workbook procedures use the same Func shape.
    Set worksheetFunctions = Application.WorksheetFunction
    Set maximum = ROneCOne.Func(worksheetFunctions, "Max") _
        .Takes(vbLong, vbLong) _
        .Returns(vbDouble)
    Set calculateTotal = ROneCOne.Func("DemoUsage.CalculateOrderTotal") _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)

    ' One notification can update the dashboard and write an audit entry.
    Set updateDashboard = ROneCOne.Action("DemoUsage.UpdateDashboard") _
        .Takes(vbString)
    Set writeAudit = ROneCOne.Action("DemoUsage.WriteAudit") _
        .Takes(vbString)
    Set notify = ROneCOne.Combine(updateDashboard, writeAudit)
    mTrace = vbNullString
    notify.Execute "Order 1042 approved"

    ' ByRef lets the delegate update the original order number in place.
    orderNumber = 1041
#If Win64 Then
    Set increment = ROneCOne.NativeAction(NextOrderNumberAddress) _
        .Takes(ROneCOne.RefOf(vbLong))
    increment.Execute ROneCOne.RefLong(orderNumber)
#End If

    ' PipeTo applies the discount first, then adds the handling charge.
    Set addHandling = price.Add(5#).AsFunc
    Set pipeline = applyDiscount.PipeTo(addHandling)

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = applyDiscount(CDbl(100))
        .Range("E7").Value2 = orderTotal(CLng(100), CLng(5))
        .Range("E8").Value2 = approvalRule(CLng(250))
        .Range("E9").Value2 = safeFalse.Run()
        .Range("E10").Value2 = maximum(CLng(4), CLng(7))
        .Range("E11").Value2 = calculateTotal(CLng(100), CLng(5))
        .Range("E12").Value2 = calculateTotal.DynamicInvoke( _
            Array(CLng(100), CLng(5)))
        .Range("E13").Value2 = mTrace
        .Range("E14").Value2 = orderNumber
        .Range("E15").Value2 = pipeline(CDbl(100))
        .Range("E16").Value2 = calculateTotal.Signature
    End With
End Sub

' -----------------------------------------------------------------------------
' Demo call targets
' -----------------------------------------------------------------------------

Public Function CalculateOrderTotal( _
    ByVal subtotal As Variant, _
    ByVal shipping As Variant _
) As Variant
    CalculateOrderTotal = subtotal + shipping
End Function

Public Sub UpdateDashboard(ByVal message As Variant)
    mTrace = "Dashboard updated"
End Sub

Public Sub WriteAudit(ByVal message As Variant)
    mTrace = mTrace & "; audit written"
End Sub

#If Win64 Then
Public Sub NextOrderNumber(ByRef orderNumber As Long)
    orderNumber = orderNumber + 1
End Sub

Public Function NextOrderNumberAddress() As LongPtr
    Dim procedureAddress As LongPtr

    CopyPointer procedureAddress, AddressOf NextOrderNumber, _
        LenB(procedureAddress)
    NextOrderNumberAddress = procedureAddress
End Function
#End If

' -----------------------------------------------------------------------------
' Benchmark and reporting
' -----------------------------------------------------------------------------

Private Sub RunDelegateBenchmark()
    Dim applyDiscount As ROneCOne
    Dim index As Long
    Dim lastResult As Variant
    Dim price As ROneCOne
    Dim started As Double

    Set price = ROneCOne.Var(vbDouble)
    Set applyDiscount = price.Multiply(0.9).AsFunc
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        lastResult = applyDiscount(CDbl(index))
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
