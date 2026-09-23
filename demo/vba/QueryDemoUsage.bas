Attribute VB_Name = "QueryDemoUsage"
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
' ROneCOne tutorial: LINQ that compiles to SQL
' ----------------------------------------------------------------------------
' This demo never touches the network. It builds one small workbook next to
' this one, queries it through the ACE OLE DB provider that ships with Office,
' and deletes it at the end, so the demo leaves nothing behind. The point is
' that the filtering happens inside the data source instead of after loading
' every row into Excel: you keep writing the same LINQ you already use, and
' the runtime turns your expression into a parameterized SQL statement.
'
' The surface mirrors what C# programmers know from IQueryable: Where,
' OrderBy, Take, Skip, SelectColumns, Count, FirstOrDefault, and ToDataTable,
' each returning a new query so a base query can be shared. ToSqlString shows
' the exact statement, which is worth looking at the first time.
'
' Two safety properties are worth watching for below. Every value you capture
' travels as a ? parameter, so a customer name containing a quote is data and
' can never become SQL. And an expression the translator cannot express
' refuses with ROneCOne.QueryError instead of quietly loading the whole table
' and filtering in memory, which would look identical until the table grew.
'
' To run it: press Alt+F8, choose RunROneCOneQueryDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 50000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneQueryDemo()
    Dim dataPath As String
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    dataPath = ThisWorkbook.Path & "\ROneCOne_Query_Demo_Data.xlsx"
    BuildDemoWorkbook dataPath
    WriteQueryExamples dataPath
    RunQueryBenchmark dataPath
    ROneCOne.File.Delete dataPath
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Len(dataPath) > 0 Then
        If ROneCOne.File.Exists(dataPath) Then ROneCOne.File.Delete dataPath
    End If
    On Error GoTo 0
    MarkDemoFailed errNumber, errorDescription
End Sub

' A closed workbook is a perfectly good database as far as ACE is concerned,
' which keeps this demo offline and free of any server.
Private Sub BuildDemoWorkbook(ByVal dataPath As String)
    Dim book As Workbook
    Dim ws As Worksheet

    On Error Resume Next
    Kill dataPath
    On Error GoTo 0

    Application.DisplayAlerts = False
    Set book = Application.Workbooks.Add
    Set ws = book.Worksheets(1)
    ws.Name = "Orders"
    ws.Range("A1:D1").Value = Array("Id", "Name", "Total", "Note")
    ws.Range("A2:D2").Value = Array(1, "Ada", 12.5, "rush")
    ws.Range("A3:D3").Value = Array(2, "Bo", 20, Empty)
    ws.Range("A4:D4").Value = Array(3, "Cy", 7.25, "gift")
    ws.Range("A5:D5").Value = Array(4, "100% done", 5, "x")
    ws.Range("A6:D6").Value = Array(5, "x' OR '1'='1", 11, "z")
    book.SaveAs dataPath, 51
    book.Close False
    Application.DisplayAlerts = True
End Sub

Private Function OpenDemoConnection(ByVal dataPath As String) As ROneCOne
    Dim conn As ROneCOne

    Set conn = ROneCOne.DbConnection( _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataPath & _
        ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
    conn.Connect
    Set OpenDemoConnection = conn
End Function

Private Sub WriteQueryExamples(ByVal dataPath As String)
    Dim baseQuery As ROneCOne
    Dim conn As ROneCOne
    Dim orders As ROneCOne
    Dim refusalTrace As String
    Dim refused As Long

    Set conn = OpenDemoConnection(dataPath)

    ' Step 1: a query is deferred. Nothing runs until a terminal like Count
    ' or ToDataTable asks for rows, so a query is cheap to build and share.
    Set orders = conn.Queryable("Orders$")

    ' Step 2: composition never mutates what it came from. baseQuery keeps
    ' its own meaning even after a narrower query is derived from it.
    Set baseQuery = orders.Where("Total").AtLeast(10)

    ' Step 3: an expression the translator cannot turn into SQL refuses
    ' loudly. A nested member path would need a join, so it raises rather
    ' than silently pulling the table into Excel and filtering there.
    refusalTrace = "unexpected acceptance"
    refused = 0
    On Error Resume Next
    Dim ignoredCount As Long
    ignoredCount = orders.Where( _
        orders.Condition("Name.Length").AtLeast(1)).Count
    refused = Err.Number
    On Error GoTo 0
    If refused = ROneCOne.QueryError Then
        refusalTrace = "refused, not scanned locally"
    End If

    ' Each line reads one result and writes it to the Examples sheet, so
    ' every behavior above shows its answer next to what the sheet expects.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = orders.Where("Total").AtLeast(10).ToSqlString
        .Range("E7").Value2 = orders.Where("Total").AtLeast(10).Count
        ' Deriving a narrower query first, then re-reading the base: the
        ' base still means what it meant.
        .Range("E8").Value2 = DerivedThenBaseCount(baseQuery)
        .Range("E9").Value2 = orders.OrderBy("Id").Take(2).ToDataTable _
            .Rows.Count
        .Range("E10").Value2 = orders.SelectColumns("Name", "Total") _
            .Take(1).ToDataTable.Columns.Count
        .Range("E11").Value2 = _
            orders.Where("Name").EqualTo("x' OR '1'='1").Count
        .Range("E12").Value2 = orders.Where("Note").EqualTo(Null).Count
        .Range("E13").Value2 = orders.Where("Name").ContainsText("100%").Count
        .Range("E14").Value2 = refusalTrace
    End With

    conn.Disconnect
End Sub

' Counting five thousand rows without moving them: the provider returns one
' number, so nothing but the answer crosses into Excel.
Private Sub RunQueryBenchmark(ByVal dataPath As String)
    Dim conn As ROneCOne
    Dim elapsed As Double
    Dim matching As Long
    Dim started As Double

    BuildBenchmarkSheet dataPath
    Set conn = OpenDemoConnection(dataPath)

    started = Timer
    matching = conn.Queryable("Bench$") _
        .Where("Value").AtLeast(BENCHMARK_ROWS / 2).Count
    elapsed = ElapsedSeconds(started)

    conn.Disconnect

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = matching
    End With
End Sub

Private Sub BuildBenchmarkSheet(ByVal dataPath As String)
    Dim book As Workbook
    Dim idx As Long
    Dim ws As Worksheet
    Dim grid() As Variant

    Application.DisplayAlerts = False
    Set book = Application.Workbooks.Open(dataPath)
    Set ws = book.Worksheets.Add
    ws.Name = "Bench"
    ws.Range("A1").Value = "Value"
    ReDim grid(1 To BENCHMARK_ROWS, 1 To 1)
    For idx = 1 To BENCHMARK_ROWS
        grid(idx, 1) = idx
    Next idx
    ws.Range("A2").Resize(BENCHMARK_ROWS, 1).Value = grid
    book.Save
    book.Close False
    Application.DisplayAlerts = True
End Sub

Private Function DerivedThenBaseCount(ByVal baseQuery As ROneCOne) As Long
    Dim narrower As ROneCOne

    Set narrower = baseQuery.Where("Total").AtMost(12.5)
    If narrower.Count = 0 Then
        DerivedThenBaseCount = -1
        Exit Function
    End If
    DerivedThenBaseCount = baseQuery.Count
End Function

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
