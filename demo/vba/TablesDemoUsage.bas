Attribute VB_Name = "TablesDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: working with Excel Tables (ListObjects)
' ----------------------------------------------------------------------------
' This demo never touches the network. It builds a real Excel Table on a
' sheet called "Sales Table", reads it, queries it, maps it to and from a
' plain class of your own, converts it to JSON and CSV and back, and writes
' results onto the sheet. The table is left in place so you can inspect it.
'
' The one thing to know before anything else: a Table is a ListObject, and a
' ListObject is not a Range. Every ROneCOne bridge takes a Range. So you
' never hand over the table itself, you hand over one of its ranges:
'
'   listObject.Range           the whole table, header row included
'   listObject.DataBodyRange   just the data rows, no header
'   listObject.ListColumns("Amount").DataBodyRange     one column of data
'
' Passing the ListObject itself raises error 438, "Object doesn't support
' this property or method", which is the single most likely first mistake.
' One example does it on purpose so you can recognize the message.
'
' After that first step everything else is ordinary work on a DataTable: a
' typed grid held in memory. Its Rows behave like any other ROneCOne
' sequence, so Where, OrderBy, GroupBy, Sum, and the rest apply unchanged.
' Nothing here requires knowing C#.
'
' To run it: press Alt+F8, choose RunROneCOneTablesDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 5000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const DATA_SHEET As String = "Sales Table"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"
Private Const TABLE_NAME As String = "Sales"
Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1

' ToObjects needs a way to make one empty instance per row, so it takes a
' delegate rather than a class name. This is that delegate's target, and it
' must be Public for ROneCOne.Func to find it by name.
Public Function NewSalesRow() As SalesRow
    Set NewSalesRow = New SalesRow
End Function

Public Sub RunROneCOneTablesDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteTableExamples BuildSalesTable()
    RunTablesBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

' Creates the Table the rest of the demo reads. Rebuilt every run so the
' demo is repeatable, and left behind so you can inspect it.
Private Function BuildSalesTable() As Object
    Dim created As Object
    Dim sheet As Worksheet

    Set sheet = EnsureSheet(DATA_SHEET)
    sheet.Cells.Clear
    On Error Resume Next
    sheet.ListObjects(TABLE_NAME).Unlist
    On Error GoTo 0
    sheet.Range("A1:C1").Value = Array("Region", "Rep", "Amount")
    sheet.Range("A2:C2").Value = Array("West", "Ada", 120)
    sheet.Range("A3:C3").Value = Array("East", "Bo", 80)
    sheet.Range("A4:C4").Value = Array("West", "Cy", 200)
    sheet.Range("A5:C5").Value = Array("North", "Dee", 45)
    sheet.Range("A6:C6").Value = Array("West", "Eve", 60)
    Set created = sheet.ListObjects.Add( _
        XL_SRC_RANGE, sheet.Range("A1:C6"), , XL_YES)
    created.Name = TABLE_NAME
    sheet.Columns("A:H").AutoFit
    Set BuildSalesTable = created
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sheet.Name = sheetName
    End If
    Set EnsureSheet = sheet
End Function

Private Sub WriteTableExamples(ByVal salesTable As Object)
    Dim Amount As Variant
    Dim amounts As ROneCOne
    Dim bodyOnly As ROneCOne
    Dim directError As Long
    Dim directTrace As String
    Dim factory As ROneCOne
    Dim fromJson As ROneCOne
    Dim grouped As ROneCOne
    Dim ignored As ROneCOne
    Dim objects As ROneCOne
    Dim rebuilt As ROneCOne
    Dim Region As Variant
    Dim Rep As Variant
    Dim rowsAfterAdd As Long
    Dim sales As ROneCOne
    Dim sheet As Worksheet
    Dim topWest As ROneCOne

    Set sheet = ThisWorkbook.Worksheets(DATA_SHEET)

    ' Step 1: find the Table by its name, then read it in one call. Passing
    ' .Range includes the header row, so the headers become column names.
    Set sales = ROneCOne.DataTableFromRange(salesTable.Range)

    ' Step 2: if you would rather skip the header row, hand over
    ' .DataBodyRange and say headers:=False. Columns are then named
    ' Column1, Column2, Column3, because there are no headers to read.
    Set bodyOnly = ROneCOne.DataTableFromRange( _
        salesTable.DataBodyRange, False)

    ' Step 3: the rows are an ordinary sequence, so every LINQ-shaped
    ' operator works on them. The bang syntax sales.Rows!Region is just a
    ' readable way to name a column when building a condition.
    Set grouped = sales.Rows.GroupBy("Region")

    ' Step 4: a DataView is the sheet-oriented alternative. It stays live
    ' against the table and knows how to write itself back to cells.
    Set topWest = ROneCOne.DataView(sales) _
        .WithFilter(sales.Rows!Region.EqualTo("West")) _
        .WithSort("Amount", True)

    ' Step 5: one column of the Table straight into a typed list.
    Set amounts = ROneCOne.ListFromRange( _
        salesTable.ListColumns("Amount").DataBodyRange)

    ' Step 6: map the rows onto instances of your own class. The property
    ' names on SalesRow match the column headers, and that is the whole
    ' mapping rule. ROneCOne.Func names the factory that makes each one.
    Set factory = ROneCOne.Func("TablesDemoUsage.NewSalesRow") _
        .Takes().Returns(vbObject)
    Set objects = sales.ToObjects(factory)

    ' Step 7: and back the other way. You list the properties to read, and
    ' each becomes a column of the new table.
    Set rebuilt = ROneCOne.DataTableFromObjects( _
        objects, Array("Region", "Rep", "Amount"), "Rebuilt")

    ' Step 8: JSON in both directions. ToJson serializes the table, a row,
    ' or a view; DeserializeTable turns an array of objects back into one.
    Set fromJson = ROneCOne.Json.DeserializeTable(sales.ToJson, "FromJson")

    ' Step 9: the mistake worth recognizing. A ListObject is not a Range,
    ' so handing the Table itself to a Range bridge raises 438. The fix is
    ' always to pass .Range or .DataBodyRange instead.
    directTrace = "unexpected success"
    On Error Resume Next
    Set ignored = ROneCOne.DataTableFromRange(salesTable)
    directError = Err.Number
    On Error GoTo 0
    If directError = 438 Then
        directTrace = "438, pass .Range instead"
    End If

    ' Step 10: write the filtered view somewhere else on the sheet. This
    ' writes headers plus the matching rows in one bulk assignment.
    topWest.ToRange sheet.Range("F1")
    sheet.Columns("F:H").AutoFit

    ' Step 11: growing the Table is Excel's job, not ROneCOne's. Add the
    ' row through the ListObject, then read the table again to pick it up.
    salesTable.ListRows.Add
    With salesTable.DataBodyRange
        .Cells(salesTable.ListRows.Count, 1).Value = "South"
        .Cells(salesTable.ListRows.Count, 2).Value = "Fay"
        .Cells(salesTable.ListRows.Count, 3).Value = 99
    End With
    rowsAfterAdd = ROneCOne.DataTableFromRange(salesTable.Range).Rows.Count

    ' Each line reads one result and writes it to the Examples sheet, so
    ' every behavior above shows its answer next to what the sheet expects.
    ' Note that sales was read before step 11 added a row, so it still
    ' reports five: a DataTable is a snapshot, not a live link to cells.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = sales.Rows.Count
        .Range("E7").Value2 = sales.Columns.Count
        .Range("E8").Value2 = CStr(sales.Rows.Item(0).Item("Rep"))
        .Range("E9").Value2 = CStr(bodyOnly.Rows.Item(0).Item("Column2"))
        .Range("E10").Value2 = sales.Rows.Count( _
            sales.Rows!Region.EqualTo("West"))
        .Range("E11").Value2 = sales.Rows.Where( _
            sales.Rows!Amount.AtLeast(100)).OrderBy("Rep").JoinText(", ", "Rep")
        .Range("E12").Value2 = sales.Rows.OrderByDescending("Amount") _
            .Take(2).Sum("Amount")
        .Range("E13").Value2 = sales.Rows.OrderBy("Region") _
            .ThenByDescending("Amount").JoinText(">", "Rep")
        .Range("E14").Value2 = sales.Rows.SelectItems("Region", vbString) _
            .Distinct.Order.JoinText(", ")
        .Range("E15").Value2 = CStr(grouped.First.Key) & " " & _
            CStr(grouped.First.Sum("Amount"))
        .Range("E16").Value2 = sales.Rows.Sum("Amount")
        .Range("E17").Value2 = sales.Rows.Average("Amount")
        .Range("E18").Value2 = sales.Rows.Max("Amount")
        .Range("E19").Value2 = CStr(sales.Rows.FirstOrDefault( _
            sales.Rows!Rep.EqualTo("Cy")).Item("Amount"))
        .Range("E20").Value2 = topWest.Count
        .Range("E21").Value2 = CStr(topWest.Item(0).Item("Rep"))
        .Range("E22").Value2 = topWest.Sum("Amount")
        .Range("E23").Value2 = amounts.Sum
        .Range("E24").Value2 = TypeName(objects.Item(0))
        .Range("E25").Value2 = CStr(objects.Item(0).Rep)
        .Range("E26").Value2 = objects.Where( _
            objects.Condition("Region").EqualTo("West")).Count
        .Range("E27").Value2 = CStr(rebuilt.Rows.Item(2).Item("Rep"))
        .Range("E28").Value2 = sales.Rows.Item(0).ToJson
        .Range("E29").Value2 = fromJson.Rows.Count
        .Range("E30").Value2 = ROneCOne.Json.DeserializeTable( _
            topWest.ToJson, "TopWest").Rows.Count
        .Range("E31").Value2 = Split(sales.ToCsv, vbCrLf)(0)
        .Range("E32").Value2 = ROneCOne.Csv.DeserializeTable( _
            sales.ToCsv, "FromCsv").Rows.Count
        .Range("E33").Value2 = directTrace
        .Range("E34").Value2 = CStr(sheet.Range("F2").Value)
        .Range("E35").Value2 = rowsAfterAdd
    End With
End Sub

' Five thousand rows out of a Table, filtered, and back to the sheet. The
' point is that each direction is one bulk call rather than a cell loop.
Private Sub RunTablesBenchmark()
    Dim Amount As Variant
    Dim benchSheet As Worksheet
    Dim benchTable As Object
    Dim elapsed As Double
    Dim index As Long
    Dim kept As ROneCOne
    Dim loaded As ROneCOne
    Dim started As Double
    Dim values() As Variant

    Set benchSheet = EnsureSheet("Benchmark Data")
    benchSheet.Cells.Clear
    On Error Resume Next
    benchSheet.ListObjects("Benchmark").Unlist
    On Error GoTo 0
    benchSheet.Range("A1:B1").Value = Array("Id", "Amount")
    ReDim values(1 To BENCHMARK_ROWS, 1 To 2)
    For index = 1 To BENCHMARK_ROWS
        values(index, 1) = index
        values(index, 2) = index * 2
    Next index
    benchSheet.Range("A2").Resize(BENCHMARK_ROWS, 2).Value = values
    Set benchTable = benchSheet.ListObjects.Add(XL_SRC_RANGE, _
        benchSheet.Range("A1").Resize(BENCHMARK_ROWS + 1, 2), , XL_YES)
    benchTable.Name = "Benchmark"

    started = Timer
    Set loaded = ROneCOne.DataTableFromRange(benchTable.Range)
    Set kept = ROneCOne.DataView(loaded) _
        .WithFilter(loaded.Rows!Amount.AtLeast(BENCHMARK_ROWS))
    kept.ToRange benchSheet.Range("E1")
    elapsed = ElapsedSeconds(started)

    benchSheet.Visible = 0
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = kept.Count
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
