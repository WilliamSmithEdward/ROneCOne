Attribute VB_Name = "DataDemoUsage"
Option Explicit

' This executable tutorial demonstrates DataTable, DataSet, DataView, and providers.

Private Const BENCHMARK_ITERATIONS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneDataDemo()
    Dim connection As ROneCOne
    Dim errorDescription As String
    Dim errorNumber As Long
    Dim fixturePath As String

    On Error GoTo DemoFailure
    fixturePath = ThisWorkbook.Path & "\ROneCOne_Data_Demo_Fixture.xlsx"
    CreateProviderFixture fixturePath
    WriteDataExamples fixturePath, connection
    RunDataBenchmark
    connection.Disconnect
    DeleteProviderFixture fixturePath
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not connection Is Nothing Then connection.Disconnect
    DeleteProviderFixture fixturePath
    On Error GoTo 0
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteDataExamples( _
    ByVal fixturePath As String, _
    ByRef connection As ROneCOne _
)
    Dim adapter As ROneCOne
    Dim child As ROneCOne
    Dim command As ROneCOne
    Dim data As ROneCOne
    Dim filled As ROneCOne
    Dim parent As ROneCOne
    Dim parentRow As ROneCOne
    Dim row As ROneCOne
    Dim scalarTask As ROneCOne
    Dim Score As Variant
    Dim table As ROneCOne
    Dim view As ROneCOne

    Set table = ROneCOne.DataTable("People")
    table.Column("Id", vbLong).AutoNumber(100&, 10&).AsUnique
    table.Column("Name", vbString).WithDefault "Unknown"
    table.Column "Score", vbLong
    Set row = table.NewRow
    row.Item("Name") = "Ada"
    row.Item("Score") = 90&
    table.AddRow row
    Set row = table.NewRow
    row.Item("Name") = "Grace"
    row.Item("Score") = 95&
    table.AddRow row
    Set view = ROneCOne.DataView(table) _
        .WithFilter(table.Rows!Score.AtLeast(90&)) _
        .WithSort("Score", True)

    Set parent = ROneCOne.DataTable("Customers")
    parent.Column("Id", vbLong).AsUnique
    Set parentRow = parent.LoadRow(Array(1&))
    Set child = ROneCOne.DataTable("Orders")
    child.Column "CustomerId", vbLong
    child.LoadRow Array(1&)
    Set data = ROneCOne.DataSet("Sales")
    data.AddTable parent
    data.AddTable child
    data.AddRelation ROneCOne.DataRelation( _
        "CustomerOrders", parent.Columns("Id"), child.Columns("CustomerId"))

    table.AcceptChanges
    table.Rows.Item(0).Item("Score") = 91&

    Set connection = ROneCOne.DbConnection( _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fixturePath & _
        ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
    connection.Connect
    Set command = ROneCOne.DbCommand( _
        "SELECT [Name], [Score] FROM [Scores$] ORDER BY [Score] DESC", _
        connection)
    Set adapter = ROneCOne.DbDataAdapter(command)
    Set filled = ROneCOne.DataTable("Scores")
    Set scalarTask = ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM [Scores$]", connection).ExecuteScalarAsync

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = table.Rows.Item(0).Item("Id")
        .Range("E7").Value2 = view.Item(0).Item("Name")
        .Range("E8").Value2 = parentRow.GetChildRows("CustomerOrders").Count
        .Range("E9").Value2 = table.GetChanges.Rows.Count
        .Range("E10").Value2 = adapter.Fill(filled)
        .Range("E11").Value2 = filled.Rows.Item(0).Item("Name")
        .Range("E12").Value2 = scalarTask.Await
    End With
End Sub

Private Sub CreateProviderFixture(ByVal fixturePath As String)
    Dim fixture As Workbook
    Dim sheet As Worksheet

    DeleteProviderFixture fixturePath
    Set fixture = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = fixture.Worksheets.Item(1)
    sheet.Name = "Scores"
    sheet.Range("A1").Value2 = "Name"
    sheet.Range("B1").Value2 = "Score"
    sheet.Range("A2").Value2 = "Ada"
    sheet.Range("B2").Value2 = 90
    sheet.Range("A3").Value2 = "Grace"
    sheet.Range("B3").Value2 = 95
    fixture.SaveAs fixturePath, xlOpenXMLWorkbook
    fixture.Close SaveChanges:=False
End Sub

Private Sub DeleteProviderFixture(ByVal fixturePath As String)
    If Len(fixturePath) = 0 Then Exit Sub
    If Len(Dir$(fixturePath)) > 0 Then Kill fixturePath
End Sub

Private Sub RunDataBenchmark()
    Dim index As Long
    Dim result As ROneCOne
    Dim started As Double
    Dim table As ROneCOne

    Set table = ROneCOne.DataTable("Benchmark")
    table.Column "Value", vbLong
    started = Timer
    For index = 1 To BENCHMARK_ITERATIONS
        table.LoadRow Array(index)
    Next index
    Set result = table.Rows.Where("Value").AtLeast(500&).ToList
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = result.Count
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
