Attribute VB_Name = "JsonDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: JSON in the spirit of System.Text.Json
' ----------------------------------------------------------------------------
' This demo never touches the network. Every document it parses is a string
' you can read a few lines below, and everything it produces stays in this
' workbook. The point is the round trip: JSON text becomes values you can
' query, tables you can put on a sheet, and objects of your own classes; all
' of them serialize back to JSON with one call.
'
' The surface mirrors what C# programmers know from System.Text.Json:
' Deserialize for a navigable tree, Serialize for compact or indented text,
' DeserializeTable to land an array of objects in a typed DataTable, and
' DeserializeInto / DeserializeObjects / ToObjects / DataTableFromObjects to
' bind against your own classes. VBA has no reflection and no Activator, so
' binding asks you for two explicit things a C# runtime would discover on its
' own: a zero-argument factory that returns a new instance, and the list of
' property names to move. Everything else is inferred.
'
' To run it: press Alt+F8, choose RunROneCOneJsonDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mTrace As String

Public Sub RunROneCOneJsonDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteJsonExamples
    RunJsonBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

' The factory the binding examples hand to the runtime. VBA cannot create an
' instance from a type name, so this one-liner stands in for Activator.
Public Function NewDemoCustomer() As DemoCustomer
    Set NewDemoCustomer = New DemoCustomer
End Function

Private Sub WriteJsonExamples()
    Dim again As ROneCOne
    Dim bindBody As String
    Dim body As String
    Dim customer As DemoCustomer
    Dim envelope As String
    Dim factory As ROneCOne
    Dim mapped As ROneCOne
    Dim orders As ROneCOne
    Dim ordersJson As String
    Dim parseError As Long
    Dim people As ROneCOne
    Dim peopleJson As String
    Dim pretty As String
    Dim tree As ROneCOne

    ' The documents. In your workbook these would arrive from a file, a cell,
    ' or an HttpClient response; here they are literals so the demo is
    ' self-contained and works without an internet connection.
    body = "{""customer"":""Ada"",""address"":{""city"":""London""}," & _
        """tags"":[""new"",""priority""]}"
    ordersJson = "[{""id"":1,""total"":19.5,""customer"":{""city"":" & _
        """London""}},{""id"":2,""total"":7.25,""customer"":{""city"":" & _
        """Paris""}},{""id"":3,""total"":12,""customer"":{""city"":" & _
        """Oslo""}}]"
    envelope = "{""data"":{""items"":[{""v"":1},{""v"":2}]}}"
    bindBody = "{""CustomerName"":""Grace"",""Age"":40,""City"":" & _
        """Cambridge""}"
    peopleJson = "[{""CustomerName"":""Ada"",""Age"":36,""City"":" & _
        """London""},{""CustomerName"":""Bo"",""Age"":50,""City"":" & _
        """Paris""}]"

    ' Deserialize returns a navigable tree: a JSON object becomes an ordered
    ' dictionary you read by name, a JSON array becomes a list you read by
    ' position, and every scalar arrives as an ordinary VBA value.
    Set tree = ROneCOne.Json.Deserialize(body)

    ' Serialize goes the other way, compact by default or indented with one
    ' flag, and an indented document parses back to the identical tree.
    pretty = ROneCOne.Json.Serialize(tree, True)
    Set again = ROneCOne.Json.Deserialize(pretty)

    ' DeserializeTable lands an array of objects in a typed DataTable. The
    ' nested customer objects become dotted columns like "customer.city",
    ' and the mixed whole and fractional totals widen to one Double column.
    Set orders = ROneCOne.Json.DeserializeTable(ordersJson, "Orders")

    ' Binding onto your own classes. DeserializeInto fills an instance you
    ' already have; DeserializeObjects builds one instance per element
    ' through the factory declared above and returns them as a typed list.
    Set customer = New DemoCustomer
    ROneCOne.Json.DeserializeInto bindBody, customer
    Set factory = ROneCOne.Func("JsonDemoUsage.NewDemoCustomer") _
        .Takes().Returns(vbObject)
    Set people = ROneCOne.Json.DeserializeObjects(peopleJson, factory)

    ' And back again: an explicit property list turns those objects into a
    ' typed DataTable, ready for ToRange or another ToJson.
    Set mapped = ROneCOne.DataTableFromObjects(people, _
        Array("CustomerName", "Age", "City"))

    ' Strictness is a feature: RFC 8259 violations such as a trailing comma
    ' raise a typed error whose message carries the offending position. Like
    ' every checked failure in the runtime, you trap it and compare against
    ' the published error number.
    mTrace = "unexpected acceptance"
    On Error Resume Next
    ROneCOne.Json.Deserialize "{""a"":1,}"
    parseError = Err.Number
    On Error GoTo 0
    If parseError = ROneCOne.JsonError Then
        mTrace = "trailing comma rejected"
    End If

    ' Every example writes its answer next to what the sheet expects.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = CStr(tree.Item("customer"))
        .Range("E7").Value2 = CStr(tree.Item("address").Item("city"))
        .Range("E8").Value2 = CStr(tree.Item("tags").Item(1))
        .Range("E9").Value2 = ROneCOne.ListOf(vbLong, 1, 2, 3).ToJson
        .Range("E10").Value2 = _
            (ROneCOne.Json.Serialize(again) = ROneCOne.Json.Serialize(tree))
        .Range("E11").Value2 = orders.Rows.Count
        .Range("E12").Value2 = CStr(orders.Rows.Item(0).Item("customer.city"))
        .Range("E13").Value2 = ROneCOne.Json.DeserializeTable( _
            envelope, "Items", "$.data.items").Rows.Count
        .Range("E14").Value2 = customer.CustomerName
        .Range("E15").Value2 = people.Item(1).CustomerName
        .Range("E16").Value2 = mapped.Rows.Count
        .Range("E17").Value2 = mTrace
    End With
End Sub

Private Sub RunJsonBenchmark()
    Dim document As String
    Dim elapsed As Double
    Dim index As Long
    Dim roundTripped As ROneCOne
    Dim started As Double
    Dim table As ROneCOne

    ' One thousand typed rows travel out to JSON text and back into a fresh
    ' typed table. The writer builds the whole document in one buffer and the
    ' reader scans it from one byte snapshot, so both directions stay flat.
    Set table = ROneCOne.DataTable("Orders")
    table.Column("Id", vbLong).AsPrimaryKey
    table.Column "Customer", vbString
    table.Column "Total", vbDouble
    For index = 1 To BENCHMARK_ROWS
        table.LoadRow Array(index, "C" & CStr(index), index * 1.5)
    Next index

    started = Timer
    document = table.ToJson
    Set roundTripped = ROneCOne.Json.DeserializeTable(document, "Orders")
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = roundTripped.Rows.Count
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
