Attribute VB_Name = "DataDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: in-memory data and providers
' ----------------------------------------------------------------------------
' This demo works with data the way a small database does, but entirely inside
' Excel. You define a table with typed columns, and every row you add is checked
' against them. You can filter and sort through a view, link two tables with a
' relationship and walk between them, track which rows changed, and even read
' another workbook through a connection, all without leaving VBA.
'
' The last part reads a second workbook as a data source. The demo creates that
' source file locally, uses it, and deletes it again, so nothing is left behind.
'
' To run it: press Alt+F8, choose RunROneCOneDataDemo, and click Run.
' ============================================================================

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

    ' Step 1: define the table's shape once. Each Column names a piece of data
    ' and its type, so every row you add later is checked against that shape.
    ' The Id column numbers itself (starting at 100, stepping by 10) and is the
    ' primary key that identifies a row; Name fills in "Unknown" when left blank.
    Set table = ROneCOne.DataTable("People")
    table.Column("Id", vbLong).AutoNumber(100&, 10&).AsPrimaryKey
    table.Column("Name", vbString).WithDefault "Unknown"
    table.Column "Score", vbLong
    table.Column "Note", vbString
    ' Add two rows. You supply Name, Score, and Note; the Id fills itself in.
    ' ROneCOne.DBNull is the data layer's way of saying "no value here" for Note.
    Set row = table.Row("Ada", 90&, ROneCOne.DBNull).Add
    Set row = table.Row("Grace", 95&, "Compiler pioneer").Add
    ' A view is a live window onto the same rows, filtered and sorted, without
    ' copying anything by hand. This one keeps scores of at least 90 and orders
    ' them high to low, so reading the view back gives the top scorer first.
    Set view = ROneCOne.DataView(table) _
        .WithFilter(table.Rows!Score.AtLeast(90&)) _
        .WithSort("Score", True)

    ' Two tables can be linked. Here every order belongs to a customer: the
    ' relationship ties the Orders CustomerId back to the Customers Id, and once
    ' linked you can start at a customer and ask for its orders (done below).
    Set parent = ROneCOne.DataTable("Customers")
    parent.Column("Id", vbLong).AsPrimaryKey
    Set parentRow = parent.LoadRow(Array(1&))
    Set child = ROneCOne.DataTable("Orders")
    child.Column "CustomerId", vbLong
    child.LoadRow Array(1&)
    Set data = ROneCOne.DataSet("Sales")
    data.AddTable parent
    data.AddTable child
    data.AddRelation ROneCOne.DataRelation( _
        "CustomerOrders", parent.Columns("Id"), child.Columns("CustomerId"))

    ' The table remembers what changed since you last saved. AcceptChanges marks
    ' everything as the new saved baseline; the edit right after it is then the
    ' only pending change, which is exactly what GetChanges reports further down.
    table.AcceptChanges
    table.Rows.Item(0).Item("Score") = 91&

    ' Now read a different workbook as if it were a database, without adding any
    ' references in the VBA editor. The connection string points at the fixture
    ' file this demo created; Connect opens it so the commands below can query it.
    Set connection = ROneCOne.DbConnection( _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fixturePath & _
        ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
    connection.Connect
    ' A command holds a query. This one asks for names and scores, highest first.
    ' An adapter runs a command and pours the results into a table you provide,
    ' so "filled" ends up holding whatever the source workbook returned.
    Set command = ROneCOne.DbCommand( _
        "SELECT [Name], [Score] FROM [Scores$] ORDER BY [Score] DESC", _
        connection)
    Set adapter = ROneCOne.DbDataAdapter(command)
    Set filled = ROneCOne.DataTable("Scores")
    ' A query that returns a single value (here, a row count) can run in the
    ' background. ExecuteScalarAsync starts it now and hands back a task; the
    ' answer is collected later with Await, once we actually need the number.
    Set scalarTask = ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM [Scores$]", connection).ExecuteScalarAsync

    ' Each line reads one result and writes it to the Examples sheet: the first
    ' row's Id, the view's top name, how many orders the customer has, how many
    ' pending changes exist, the rows the adapter filled, the counted total, and
    ' so on, so every feature above shows its answer next to the others.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = table.Rows.Item(0).Item("Id")
        .Range("E7").Value2 = view.Item(0).Item("Name")
        .Range("E8").Value2 = parentRow.GetChildRows("CustomerOrders").Count
        .Range("E9").Value2 = table.GetChanges.Rows.Count
        .Range("E10").Value2 = adapter.Fill(filled)
        .Range("E11").Value2 = filled.Rows.Item(0).Item("Name")
        .Range("E12").Value2 = scalarTask.Await
        .Range("E13").Value2 = IsNull(table.Rows.Item(0).Item("Note"))
        .Range("E14").Value2 = connection.AsyncMode
        .Range("E15").Value2 = connection.SupportsNativeAsync
    End With
End Sub

Private Sub CreateProviderFixture(ByVal fixturePath As String)
    ' Build the throwaway source workbook the connection reads from. It is a
    ' plain sheet named "Scores" with two names and two numbers, saved next to
    ' this workbook and deleted at the end so the demo leaves nothing behind.
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

    ' Load a thousand rows, then filter them, and time the whole thing. This
    ' shows the in-memory table stays fast enough for everyday workbook data.
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
