Attribute VB_Name = "DemoUsage"
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
' ROneCOne tutorial: delegates and expressions
' ----------------------------------------------------------------------------
' A "delegate" is just a piece of work saved in a variable. Instead of running
' a calculation right away, you describe it once, keep it, and run it whenever
' you like, as many times as you like, passing different inputs each time.
'
' This demo builds a few small business rules that way: a discount, an order
' total, an approval check, and a notification that fans out to two places. You
' do not need to understand the internals to follow along. Read top to bottom;
' every example names the useful result first, then shows how it was built.
'
' To run it: press Alt+F8, choose RunROneCOneDemo, and click Run. The results
' appear on the "Examples" worksheet, one per row.
' ============================================================================

#If Win64 Then
Private Declare PtrSafe Sub CopyPointer Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As LongPtr, _
    ByRef src As LongPtr, _
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
    Dim errNumber As Long

    On Error GoTo DemoFailure

    WriteDelegateExamples
    RunDelegateBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

' -----------------------------------------------------------------------------
' Practical delegate examples
' -----------------------------------------------------------------------------

Private Sub WriteDelegateExamples()
    Dim addHandling As ROneCOne
    Dim orderAmount As ROneCOne
    Dim applyDiscount As ROneCOne
    Dim approvalRule As ROneCOne
    Dim calculateTotal As ROneCOne
    Dim addOne As ROneCOne
    Dim largest As ROneCOne
    Dim announce As ROneCOne
    Dim orderNumber As Long
    Dim orderTotal As ROneCOne
    Dim pipeline As ROneCOne
    Dim listPrice As ROneCOne
    Dim safeFalse As ROneCOne
    Dim shipping As ROneCOne
    Dim dashboard As ROneCOne
    Dim worksheetFunctions As Object
    Dim audit As ROneCOne

    ' Step 1: describe placeholders for the numbers a rule will receive later.
    ' "Var" means "a value of this type that I will supply when I run the rule."
    ' Think of them as the blanks in a fill-in-the-blank formula.
    Set orderAmount = ROneCOne.Var(vbLong)
    Set shipping = ROneCOne.Var(vbLong)
    Set listPrice = ROneCOne.Var(vbDouble)

    ' Now write the formulas over those blanks and freeze each one into a
    ' runnable rule with AsFunc. applyDiscount takes a price off ten percent;
    ' orderTotal adds shipping to an amount; approvalRule is true only when the
    ' amount is at least 100 and below 1000. Nothing runs yet; these are recipes.
    Set applyDiscount = listPrice.Multiply(0.9).AsFunc
    Set orderTotal = orderAmount.Add(shipping).AsFunc
    Set approvalRule = orderAmount.AtLeast(100) _
        .AndAlso(orderAmount.LessThan(1000)) _
        .AsFunc

    ' AndAlso is a "short-circuit" and: once the left side is False, the right
    ' side is never looked at. Here the left side is False, so the deliberately
    ' unsafe divide-by-zero on the right is never reached and cannot fail.
    Set safeFalse = ROneCOne.Value(False) _
        .AndAlso(ROneCOne.Value(1).Divide(0)) _
        .AsFunc

    ' A rule can also wrap work that already exists. Point Func at a built-in
    ' Excel worksheet function, or at one of your own workbook procedures by
    ' name. Takes and Returns state the expected argument and result types so a
    ' wrong call is caught before the target ever runs.
    Set worksheetFunctions = Application.WorksheetFunction
    Set largest = ROneCOne.Func(worksheetFunctions, "Max") _
        .Takes(vbLong, vbLong) _
        .Returns(vbDouble)
    Set calculateTotal = ROneCOne.Func("DemoUsage.CalculateOrderTotal") _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)

    ' Combine joins several actions into one. Running "notify" once updates the
    ' dashboard and writes the audit entry, in that order, from a single call.
    Set dashboard = ROneCOne.Action("DemoUsage.UpdateDashboard") _
        .Takes(vbString)
    Set audit = ROneCOne.Action("DemoUsage.WriteAudit") _
        .Takes(vbString)
    Set announce = ROneCOne.Combine(dashboard, audit)
    mTrace = vbNullString
    announce.Execute "Order 1042 approved"

    ' Most rules only read their inputs. This one changes a variable in place:
    ' passing orderNumber "by reference" lets the native action bump it from
    ' 1041 to 1042. This is the one place a delegate writes back to your data.
    orderNumber = 1041
#If Win64 Then
    Set addOne = ROneCOne.NativeAction(NextOrderNumberAddress) _
        .Takes(ROneCOne.RefOf(vbLong))
    addOne.Execute ROneCOne.RefLong(orderNumber)
#End If

    ' PipeTo chains two rules so one feeds the next: apply the discount, then
    ' add a handling charge to that discounted price. The output of the first
    ' becomes the input of the second.
    Set addHandling = listPrice.Add(5#).AsFunc
    Set pipeline = applyDiscount.PipeTo(addHandling)

    ' Everything above only described work. This is where the rules finally run.
    ' Calling a rule looks like calling a function: applyDiscount(100) runs the
    ' saved recipe with 100 as the input. Each line writes one answer to the
    ' Examples sheet so you can see the result next to the rule that produced it.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = applyDiscount(100)
        .Range("E7").Value2 = orderTotal(100, 5)
        .Range("E8").Value2 = approvalRule(250)
        .Range("E9").Value2 = safeFalse.Run()
        .Range("E10").Value2 = largest(4, 7)
        .Range("E11").Value2 = calculateTotal(100, 5)
        .Range("E12").Value2 = calculateTotal.DynamicInvoke( _
            Array(100, 5))
        .Range("E13").Value2 = mTrace
        .Range("E14").Value2 = orderNumber
        .Range("E15").Value2 = pipeline(100)
        .Range("E16").Value2 = calculateTotal.Signature
    End With
End Sub

' -----------------------------------------------------------------------------
' Demo call targets
' ----------------------------------------------------------------------------
' These are ordinary workbook procedures. The rules above reach them by name,
' which shows that ROneCOne wraps the plain VBA code you already write; there is
' nothing special about the procedures themselves.
' -----------------------------------------------------------------------------

Public Function CalculateOrderTotal( _
    ByVal itemsTotal As Variant, _
    ByVal shipping As Variant _
) As Variant
    CalculateOrderTotal = itemsTotal + shipping
End Function

Public Sub UpdateDashboard(ByVal msg As Variant)
    mTrace = "Dashboard updated"
End Sub

Public Sub WriteAudit(ByVal msg As Variant)
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
    Dim idx As Long
    Dim lastResult As Variant
    Dim listPrice As ROneCOne
    Dim started As Double

    ' Build one rule, then run it ten thousand times and time the loop. This
    ' shows the per-call cost is small enough for everyday workbook use.
    Set listPrice = ROneCOne.Var(vbDouble)
    Set applyDiscount = listPrice.Multiply(0.9).AsFunc
    started = Timer
    For idx = 1 To BENCHMARK_ITERATIONS
        lastResult = applyDiscount(idx)
    Next idx

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
