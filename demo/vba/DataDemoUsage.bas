Attribute VB_Name = "DataDemoUsage"
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
    Dim conn As ROneCOne
    Dim errorDescription As String
    Dim errNumber As Long
    Dim fixturePath As String

    On Error GoTo DemoFailure
    fixturePath = ThisWorkbook.Path & "\ROneCOne_Data_Demo_Fixture.xlsx"
    CreateProviderFixture fixturePath
    WriteDataExamples fixturePath, conn
    RunDataBenchmark
    conn.Disconnect
    DeleteProviderFixture fixturePath
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not conn Is Nothing Then conn.Disconnect
    DeleteProviderFixture fixturePath
    On Error GoTo 0
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteDataExamples( _
    ByVal fixturePath As String, _
    ByRef conn As ROneCOne _
)
    Dim adapter As ROneCOne
    Dim orders As ROneCOne
    Dim scoresQuery As ROneCOne
    Dim sales As ROneCOne
    Dim filled As ROneCOne
    Dim customers As ROneCOne
    Dim parentRow As ROneCOne
    Dim person As ROneCOne
    Dim scalarTask As ROneCOne
    Dim Score As Variant
    Dim people As ROneCOne
    Dim topScorers As ROneCOne

    ' Step 1: define the table's shape once. Each Column names a piece of data
    ' and its type, so every row you add later is checked against that shape.
    ' The Id column numbers itself (starting at 100, stepping by 10) and is the
    ' primary key that identifies a row; Name fills in "Unknown" when left blank.
    Set people = ROneCOne.DataTable("People")
    people.Column("Id", vbLong).AutoNumber(100, 10).AsPrimaryKey
    people.Column("Name", vbString).WithDefault "Unknown"
    people.Column "Score", vbLong
    people.Column "Note", vbString
    ' Add two rows. You supply Name, Score, and Note; the Id fills itself in.
    ' ROneCOne.DBNull is the data layer's way of saying "no value here" for Note.
    Set person = people.Row("Ada", 90, ROneCOne.DBNull).Add
    Set person = people.Row("Grace", 95, "Compiler pioneer").Add
    ' A view is a live window onto the same rows, filtered and sorted, without
    ' copying anything by hand. This one keeps scores of at least 90 and orders
    ' them high to low, so reading the view back gives the top scorer first.
    Set topScorers = ROneCOne.DataView(people) _
        .WithFilter(people.Rows!Score.AtLeast(90)) _
        .WithSort("Score", True)

    ' Two tables can be linked. Here every order belongs to a customer: the
    ' relationship ties the Orders CustomerId back to the Customers Id, and once
    ' linked you can start at a customer and ask for its orders (done below).
    Set customers = ROneCOne.DataTable("Customers")
    customers.Column("Id", vbLong).AsPrimaryKey
    Set parentRow = customers.LoadRow(Array(1))
    Set orders = ROneCOne.DataTable("Orders")
    orders.Column "CustomerId", vbLong
    orders.LoadRow Array(1)
    Set sales = ROneCOne.DataSet("Sales")
    sales.AddTable customers
    sales.AddTable orders
    sales.AddRelation ROneCOne.DataRelation( _
        "CustomerOrders", customers.Columns("Id"), orders.Columns("CustomerId"))

    ' The table remembers what changed since you last saved. AcceptChanges marks
    ' everything as the new saved baseline; the edit right after it is then the
    ' only pending change, which is exactly what GetChanges reports further down.
    people.AcceptChanges
    people.Rows.Item(0).Item("Score") = 91

    ' Now read a different workbook as if it were a database, without adding any
    ' references in the VBA editor. The connection string points at the fixture
    ' file this demo created; Connect opens it so the commands below can query it.
    Set conn = ROneCOne.DbConnection( _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fixturePath & _
        ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";")
    conn.Connect
    ' A command holds a query. This one asks for names and scores, highest first.
    ' An adapter runs a command and pours the results into a table you provide,
    ' so "filled" ends up holding whatever the source workbook returned.
    Set scoresQuery = ROneCOne.DbCommand( _
        "SELECT [Name], [Score] FROM [Scores$] ORDER BY [Score] DESC", _
        conn)
    Set adapter = ROneCOne.DbDataAdapter(scoresQuery)
    Set filled = ROneCOne.DataTable("Scores")
    ' A query that returns a single value (here, a row count) can run in the
    ' background. ExecuteScalarAsync starts it now and hands back a task; the
    ' answer is collected later with Await, once we actually need the number.
    Set scalarTask = ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM [Scores$]", conn).ExecuteScalarAsync

    ' Each line reads one result and writes it to the Examples sheet: the first
    ' row's Id, the view's top name, how many orders the customer has, how many
    ' pending changes exist, the rows the adapter filled, the counted total, and
    ' so on, so every feature above shows its answer next to the others.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = people.Rows.Item(0).Item("Id")
        .Range("E7").Value2 = topScorers.Item(0).Item("Name")
        .Range("E8").Value2 = parentRow.GetChildRows("CustomerOrders").Count
        .Range("E9").Value2 = people.GetChanges.Rows.Count
        .Range("E10").Value2 = adapter.Fill(filled)
        .Range("E11").Value2 = filled.Rows.Item(0).Item("Name")
        .Range("E12").Value2 = scalarTask.Await
        .Range("E13").Value2 = IsNull(people.Rows.Item(0).Item("Note"))
        .Range("E14").Value2 = conn.AsyncMode
        .Range("E15").Value2 = conn.State
    End With
End Sub

Private Sub CreateProviderFixture(ByVal fixturePath As String)
    ' Build the throwaway source workbook the connection reads from. It is a
    ' plain sheet named "Scores" with two names and two numbers, saved next to
    ' this workbook and deleted at the end so the demo leaves nothing behind.
    Dim fixture As Workbook
    Dim ws As Worksheet

    DeleteProviderFixture fixturePath
    Set fixture = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = fixture.Worksheets.Item(1)
    ws.Name = "Scores"
    ws.Range("A1").Value2 = "Name"
    ws.Range("B1").Value2 = "Score"
    ws.Range("A2").Value2 = "Ada"
    ws.Range("B2").Value2 = 90
    ws.Range("A3").Value2 = "Grace"
    ws.Range("B3").Value2 = 95
    fixture.SaveAs fixturePath, xlOpenXMLWorkbook
    fixture.Close SaveChanges:=False
End Sub

Private Sub DeleteProviderFixture(ByVal fixturePath As String)
    If Len(fixturePath) = 0 Then Exit Sub
    If Len(Dir$(fixturePath)) > 0 Then Kill fixturePath
End Sub

Private Sub RunDataBenchmark()
    Dim idx As Long
    Dim outcome As ROneCOne
    Dim started As Double
    Dim benchTable As ROneCOne

    ' Load a thousand rows, then filter them, and time the whole thing. This
    ' shows the in-memory table stays fast enough for everyday workbook data.
    Set benchTable = ROneCOne.DataTable("Benchmark")
    benchTable.Column "Value", vbLong
    started = Timer
    For idx = 1 To BENCHMARK_ITERATIONS
        benchTable.LoadRow Array(idx)
    Next idx
    Set outcome = benchTable.Rows.Where("Value").AtLeast(500).ToList
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ITERATIONS
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = outcome.Count
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
