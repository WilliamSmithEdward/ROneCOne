Attribute VB_Name = "ListObjectDemoUsage"
Option Explicit

' ROneCOne 1.9.0, released 2026-08-02
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
' ROneCOne tutorial: Excel Tables, queried and written back in place
' ----------------------------------------------------------------------------
' This demo never touches the network. It builds one Excel Table called
' "Sales" on the "Sales Table" sheet and works only on that, leaving it there
' at the end so the write-back is visible. The point is that the ListObject is
' the input: ROneCOne.Table takes the Table rather than one of its ranges, and
' the table it returns remembers where it came from, so it can re-read that
' sheet and resize it to fit whatever rows you hand back.
'
' The surface mirrors what C# programmers know from System.Data: a DataTable
' whose Rows you can query, and a DataView that filters and sorts without
' copying. Rows is an ordinary sequence, so Where, OrderBy, GroupBy, and the
' aggregates apply unchanged, and ToObjects maps rows onto instances of your
' own class.
'
' One property is worth watching for below. WriteBack owns the Table's extent:
' fewer rows shrinks it and clears the cells it gave up, more rows grows it,
' and a totals row survives either way. Excel does none of that on its own.
'
' To run it: press Alt+F8, choose RunROneCOneListObjectDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 5000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const DATA_SHEET As String = "Sales Table"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"
Private Const TABLE_NAME As String = "Sales"
Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1

' ToObjects makes one instance per row, so it takes a delegate rather than a
' class name. This is that delegate's target, and it must be Public for
' ROneCOne.Func to resolve it by name.
Public Function NewSalesRow() As SalesRow
    Set NewSalesRow = New SalesRow
End Function

Public Sub RunROneCOneListObjectDemo()
    Dim errorDescription As String
    Dim errNumber As Long
    Dim salesTable As Object

    On Error GoTo DemoFailure
    Set salesTable = BuildSalesTable()
    WriteReadingExamples salesTable
    WriteQueryExamples salesTable
    WriteMappingExamples salesTable
    WriteWriteBackExamples salesTable
    RunListObjectBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

' Creates the Table the rest of the demo works on. Rebuilt every run so the
' demo is repeatable however the last run left it.
Private Function BuildSalesTable() As Object
    Dim created As Object
    Dim ws As Worksheet

    Set ws = EnsureSheet(DATA_SHEET)
    ws.Cells.Clear
    On Error Resume Next
    ws.ListObjects(TABLE_NAME).Unlist
    On Error GoTo 0
    ws.Range("A1:C1").Value = Array("Region", "Rep", "Amount")
    ws.Range("A2:C2").Value = Array("West", "Ada", 120)
    ws.Range("A3:C3").Value = Array("East", "Bo", 80)
    ws.Range("A4:C4").Value = Array("West", "Cy", 200)
    ws.Range("A5:C5").Value = Array("North", "Dee", 45)
    ws.Range("A6:C6").Value = Array("West", "Eve", 60)
    Set created = ws.ListObjects.Add( _
        XL_SRC_RANGE, ws.Range("A1:C6"), , XL_YES)
    created.Name = TABLE_NAME
    created.TableStyle = "TableStyleMedium2"
    ws.Columns("A:C").AutoFit
    Set BuildSalesTable = created
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

' --- Reading -----------------------------------------------------------------
Private Sub WriteReadingExamples(ByVal salesTable As Object)
    Dim attached As ROneCOne
    Dim bodyOnly As ROneCOne
    Dim direct As ROneCOne
    Dim oneColumn As ROneCOne
    Dim typed As ROneCOne

    ' The Table itself is the argument, and this is the form that stays
    ' attached to it.
    Set attached = ROneCOne.Table(salesTable)

    ' The plain bridge takes it too, when you do not need the attachment.
    Set direct = ROneCOne.DataTableFromRange(salesTable)

    ' Ask for the body instead of the whole table and the header row is left
    ' out, so the columns are named Column1, Column2, Column3.
    Set bodyOnly = ROneCOne.DataTableFromRange(salesTable, False)

    ' A single table column reads into a typed list.
    Set oneColumn = ROneCOne.ListFromRange(salesTable.ListColumns("Amount"))

    ' Or load into a table whose columns you declared yourself.
    Set typed = ROneCOne.DataTable("Typed")
    typed.Column "Region", vbVariant
    typed.Column "Rep", vbVariant
    typed.Column "Amount", vbVariant
    typed.LoadFromRange salesTable

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = attached.Rows.Count
        .Range("E7").Value2 = attached.Columns.Count
        .Range("E8").Value2 = attached.TableName
        .Range("E9").Value2 = CStr(attached.Rows.Item(0).Item("Rep"))
        .Range("E10").Value2 = direct.Rows.Count
        .Range("E11").Value2 = CStr(bodyOnly.Rows.Item(0).Item("Column2"))
        .Range("E12").Value2 = oneColumn.Sum
        .Range("E13").Value2 = typed.Rows.Count
    End With
End Sub

' --- Querying ----------------------------------------------------------------
Private Sub WriteQueryExamples(ByVal salesTable As Object)
    Dim Amount As Variant
    Dim grouped As ROneCOne
    Dim Region As Variant
    Dim Rep As Variant
    Dim sales As ROneCOne
    Dim topWest As ROneCOne

    Set sales = ROneCOne.Table(salesTable)

    ' Rows is a sequence, so every LINQ-shaped operator applies to it. The
    ' bang syntax sales.Rows!Region is a readable way to name a column when
    ' building a condition; it means the same as .Condition("Region").
    Set grouped = sales.Rows.GroupBy("Region")

    ' A DataView is the sheet-oriented alternative: a live filtered and
    ' sorted window that knows how to write itself back.
    Set topWest = ROneCOne.DataView(sales) _
        .WithFilter(sales.Rows!Region.EqualTo("West")) _
        .WithSort("Amount", True)

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E14").Value2 = sales.Rows.Count( _
            sales.Rows!Region.EqualTo("West"))
        .Range("E15").Value2 = sales.Rows.Where( _
            sales.Rows!Amount.AtLeast(100)).OrderBy("Rep").JoinText(", ", "Rep")
        .Range("E16").Value2 = sales.Rows.OrderByDescending("Amount") _
            .Take(2).Sum("Amount")
        .Range("E17").Value2 = sales.Rows.OrderBy("Region") _
            .ThenByDescending("Amount").JoinText(">", "Rep")
        .Range("E18").Value2 = sales.Rows.SelectItems("Region", vbString) _
            .Distinct.Order.JoinText(", ")
        .Range("E19").Value2 = CStr(grouped.First.Key) & " " & _
            CStr(grouped.First.Sum("Amount"))
        .Range("E20").Value2 = grouped.Count
        .Range("E21").Value2 = sales.Rows.Sum("Amount")
        .Range("E22").Value2 = sales.Rows.Average("Amount")
        .Range("E23").Value2 = sales.Rows.Max("Amount")
        .Range("E24").Value2 = CStr(sales.Rows.FirstOrDefault( _
            sales.Rows!Rep.EqualTo("Cy")).Item("Amount"))
        .Range("E25").Value2 = sales.Rows.AnyItem( _
            sales.Rows!Amount.AtLeast(150))
        .Range("E26").Value2 = topWest.Count
        .Range("E27").Value2 = CStr(topWest.Item(0).Item("Rep"))
        .Range("E28").Value2 = topWest.Sum("Amount")
    End With
End Sub

' --- Mapping to your own objects, and to text --------------------------------
Private Sub WriteMappingExamples(ByVal salesTable As Object)
    Dim factory As ROneCOne
    Dim objects As ROneCOne
    Dim rebuilt As ROneCOne
    Dim Region As Variant
    Dim sales As ROneCOne
    Dim topWest As ROneCOne

    Set sales = ROneCOne.Table(salesTable)

    ' The property names on SalesRow match the column headers, and that match
    ' is the entire mapping rule. ROneCOne.Func names the factory that makes
    ' each instance.
    Set factory = ROneCOne.Func("ListObjectDemoUsage.NewSalesRow") _
        .Takes().Returns(vbObject)
    Set objects = sales.ToObjects(factory)

    ' And back the other way: you list the properties to read, and each one
    ' becomes a column of the new table.
    Set rebuilt = ROneCOne.DataTableFromObjects( _
        objects, Array("Region", "Rep", "Amount"), "Rebuilt")

    Set topWest = ROneCOne.DataView(sales) _
        .WithFilter(sales.Rows!Region.EqualTo("West"))

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E29").Value2 = TypeName(objects.Item(0))
        .Range("E30").Value2 = CStr(objects.Item(0).Rep)
        .Range("E31").Value2 = objects.Item(2).Amount
        .Range("E32").Value2 = objects.Where( _
            objects.Condition("Region").EqualTo("West")).Count
        .Range("E33").Value2 = rebuilt.Rows.Count
        .Range("E34").Value2 = CStr(rebuilt.Rows.Item(2).Item("Rep"))
        .Range("E35").Value2 = CBool(rebuilt.ToJson = sales.ToJson)
        .Range("E36").Value2 = sales.Rows.Item(0).ToJson
        .Range("E37").Value2 = ROneCOne.Json.DeserializeTable( _
            sales.ToJson, "FromJson").Rows.Count
        .Range("E38").Value2 = ROneCOne.Json.DeserializeTable( _
            topWest.ToJson, "TopWest").Rows.Count
        .Range("E39").Value2 = Split(sales.ToCsv, vbCrLf)(0)
        .Range("E40").Value2 = ROneCOne.Csv.DeserializeTable( _
            sales.ToCsv, "FromCsv").Rows.Count
    End With
End Sub

' --- Writing back into the Table itself --------------------------------------
Private Sub WriteWriteBackExamples(ByVal salesTable As Object)
    Dim afterGrow As Long
    Dim afterShrink As Long
    Dim Region As Variant
    Dim sales As ROneCOne
    Dim ws As Worksheet
    Dim topWest As ROneCOne
    Dim vacatedCleared As Boolean
    Dim writtenBack As Long

    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    Set sales = ROneCOne.Table(salesTable)

    ' Keep only the West rows, biggest first, and push that into the Table.
    ' The Table shrinks from five rows to three, and the two rows it gave up
    ' are cleared rather than left sitting under the table.
    Set topWest = ROneCOne.DataView(sales) _
        .WithFilter(sales.Rows!Region.EqualTo("West")) _
        .WithSort("Amount", True)
    writtenBack = sales.WriteBack(topWest)
    afterShrink = salesTable.ListRows.Count
    vacatedCleared = IsEmpty(ws.Range("A6").Value)

    ' Refresh re-reads the sheet in place, so the same variable now sees the
    ' three rows that are actually there.
    sales.Refresh

    ' Add two rows in memory and write again: the Table grows to match.
    sales.LoadRow Array("North", "Dee", 45)
    sales.LoadRow Array("South", "Fay", 99)
    sales.WriteBack
    afterGrow = salesTable.ListRows.Count

    ' A totals row is not data. It stays out of Rows, and it survives a
    ' write back even though resizing a table has to switch it off first.
    salesTable.ShowTotals = True
    Set sales = ROneCOne.Table(salesTable)

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E41").Value2 = writtenBack
        .Range("E42").Value2 = afterShrink
        .Range("E43").Value2 = vacatedCleared
        .Range("E44").Value2 = CStr(ws.Range("B2").Value)
        .Range("E45").Value2 = afterGrow
        .Range("E46").Value2 = CStr(ws.Range("B6").Value)
        .Range("E47").Value2 = sales.Rows.Count
        .Range("E48").Value2 = sales.WriteBack
        .Range("E49").Value2 = CBool(salesTable.ShowTotals)
    End With

    salesTable.ShowTotals = False
    ws.Columns("A:C").AutoFit
End Sub

' Five thousand rows out of a real Table, filtered, and written straight back
' into it. Each direction is one bulk call, and the Table is resized to fit
' rather than overwritten in place.
Private Sub RunListObjectBenchmark()
    Dim benchSheet As Worksheet
    Dim benchTable As Object
    Dim elapsed As Double
    Dim idx As Long
    Dim kept As ROneCOne
    Dim loaded As ROneCOne
    Dim started As Double
    Dim grid() As Variant

    Set benchSheet = EnsureSheet("Benchmark Data")
    benchSheet.Cells.Clear
    On Error Resume Next
    benchSheet.ListObjects("Benchmark").Unlist
    On Error GoTo 0
    benchSheet.Range("A1:B1").Value = Array("Id", "Amount")
    ReDim grid(1 To BENCHMARK_ROWS, 1 To 2)
    For idx = 1 To BENCHMARK_ROWS
        grid(idx, 1) = idx
        grid(idx, 2) = idx * 2
    Next idx
    benchSheet.Range("A2").Resize(BENCHMARK_ROWS, 2).Value = grid
    Set benchTable = benchSheet.ListObjects.Add(XL_SRC_RANGE, _
        benchSheet.Range("A1").Resize(BENCHMARK_ROWS + 1, 2), , XL_YES)
    benchTable.Name = "Benchmark"

    started = Timer
    Set loaded = ROneCOne.Table(benchTable)
    Set kept = ROneCOne.DataView(loaded) _
        .WithFilter(loaded.Rows.Condition("Amount").AtLeast(BENCHMARK_ROWS))
    loaded.WriteBack kept
    elapsed = ElapsedSeconds(started)

    benchSheet.Visible = 0
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = benchTable.ListRows.Count
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
