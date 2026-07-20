Attribute VB_Name = "ExceptionsDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: structured Try, Catch, and Finally
' ----------------------------------------------------------------------------
' Real work can fail, and some cleanup must happen no matter what. This pattern
' makes both explicit in one place: TRY the work, CATCH a specific failure so
' you can recover from it, and FINALLY run cleanup that always happens, whether
' the work succeeded, failed and was caught, or failed and was not.
'
' The example imports sales rows and always closes the file afterwards. A Catch
' stands ready for one known problem (a bad amount); because this safe demo
' feeds a valid row, the Catch is never needed and the import simply succeeds.
'
' To run it: press Alt+F8, choose RunROneCOneExceptionsDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ITERATIONS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const INVALID_AMOUNT_ERROR As Long = vbObjectError + 1101
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
    Dim closeImportFile As ROneCOne
    Dim importAttempt As ROneCOne
    Dim importWork As ROneCOne
    Dim skipBadRow As ROneCOne
    Dim successfulImport As ROneCOne

    ' Name the three steps as ordinary procedures. The recovery step,
    ' skipBadRow, is allowed one argument (ROneCOne.Exception): if the import
    ' fails, ROneCOne hands it the captured error so recovery can be specific.
    Set importWork = ROneCOne.Action( _
        "ExceptionsDemoUsage.ImportSales").Takes()
    Set skipBadRow = ROneCOne.Action( _
        "ExceptionsDemoUsage.SkipBadRow").Takes(ROneCOne.Exception)
    Set closeImportFile = ROneCOne.Action( _
        "ExceptionsDemoUsage.CloseImportFile").Takes()

    ' Assemble the protected operation and then run it with Execute. Read it as
    ' a sentence: try the import; if a bad-amount error happens, skip that row;
    ' and no matter what, close the file at the end.
    mTrace = vbNullString
    Set importAttempt = ROneCOne.Try(importWork) _
        .Catch(INVALID_AMOUNT_ERROR, skipBadRow) _
        .Finally(closeImportFile)
    importAttempt.Execute
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = mTrace
        .Range("E7").Value2 = Not importAttempt Is Nothing
        .Range("E8").Value2 = "No import error"
    End With

    ' Finally is not only for failures. Here the import succeeds and the very
    ' same cleanup still runs, which is why files never leak open in either case.
    Set successfulImport = ROneCOne.Action( _
        "ExceptionsDemoUsage.ImportValidSales").Takes()
    Set importAttempt = ROneCOne.Try(successfulImport) _
        .Finally(closeImportFile)
    mTrace = vbNullString
    importAttempt.Execute
    ThisWorkbook.Worksheets(EXAMPLES_SHEET).Range("E9").Value2 = mTrace
End Sub

Public Sub ImportSales()
    AppendTrace "3 rows imported"
End Sub

Public Sub SkipBadRow(ByVal errorInfo As Variant)
    ' Catch passes the original VBA error here so recovery can be specific.
    If errorInfo.ErrorNumber = INVALID_AMOUNT_ERROR Then
        AppendTrace "Row 7 skipped"
    End If
End Sub

Public Sub CloseImportFile()
    AppendTrace "file closed"
End Sub

Public Sub ImportValidSales()
    AppendTrace "3 rows imported"
End Sub

Private Sub AppendTrace(ByVal message As String)
    If Len(mTrace) > 0 Then mTrace = mTrace & "; "
    mTrace = mTrace & message
End Sub

Private Sub RunExceptionBenchmark()
    Dim importAttempt As ROneCOne
    Dim index As Long
    Dim started As Double
    Dim successfulImport As ROneCOne

    Set successfulImport = ROneCOne.Action( _
        "ExceptionsDemoUsage.ImportValidSales").Takes()
    Set importAttempt = ROneCOne.Try(successfulImport)
    mTrace = vbNullString
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        importAttempt.Execute
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
