Attribute VB_Name = "TestTasksAndData"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mCurrentTest As String

' The live suite machine provisions a local SQL Server default instance.
' Integrated security keeps every credential out of the repository.
Private Const SQL_SERVER_CONNECTION As String = _
    "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=tempdb;" & _
    "Integrated Security=SSPI;"

Public Sub RunROneCOneTaskAndDataTests()
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    On Error GoTo FatalFailure
    ResetResults

    mCurrentTest = "TestTaskLifecycle"
    TestTaskLifecycle
    mCurrentTest = "TestTaskCombinators"
    TestTaskCombinators
    mCurrentTest = "TestCancellation"
    TestCancellation
    mCurrentTest = "TestTaskCompletionAndProgress"
    TestTaskCompletionAndProgress
    mCurrentTest = "TestTaskTimeoutAndContinuation"
    TestTaskTimeoutAndContinuation
    mCurrentTest = "TestDataTable"
    TestDataTable
    mCurrentTest = "TestDataSetRelations"
    TestDataSetRelations
    mCurrentTest = "TestCompositeRelations"
    TestCompositeRelations
    mCurrentTest = "TestAdvancedDataTable"
    TestAdvancedDataTable
    mCurrentTest = "TestDataViewAndMerge"
    TestDataViewAndMerge
    mCurrentTest = "TestSnapshotCachingAndViewRefresh"
    TestSnapshotCachingAndViewRefresh
    mCurrentTest = "TestConstraintMaintenance"
    TestConstraintMaintenance
    mCurrentTest = "TestJsonSurface"
    TestJsonSurface
    mCurrentTest = "TestRangeBridge"
    TestRangeBridge
    mCurrentTest = "TestRelationConstraints"
    TestRelationConstraints
    mCurrentTest = "TestExistingRelationValidation"
    TestExistingRelationValidation
    mCurrentTest = "TestProviderSurface"
    TestProviderSurface
    mCurrentTest = "TestSqlServerProvider"
    TestSqlServerProvider
    mCurrentTest = "TestFileSystemSurface"
    TestFileSystemSurface
    mCurrentTest = "TestCsvSurface"
    TestCsvSurface
    mCurrentTest = "TestProcessSurface"
    TestProcessSurface
    mCurrentTest = "TestRegexSurface"
    TestRegexSurface
    mCurrentTest = "TestHashSurface"
    TestHashSurface
    mCurrentTest = "TestDateTimeSurface"
    TestDateTimeSurface
    mCurrentTest = "TestStringsSurface"
    TestStringsSurface
    mCurrentTest = "TestEscapingSurface"
    TestEscapingSurface
    mCurrentTest = "TestXmlSurface"
    TestXmlSurface
    mCurrentTest = "TestHttpSurface"
    TestHttpSurface
    mCurrentTest = vbNullString

    With ThisWorkbook.Worksheets("Task and Data Results")
        .Range("B2").Value2 = mPassed
        .Range("B3").Value2 = mFailed
        .Range("B4").Value2 = IIf(mFailed = 0, "PASS", "FAIL")
    End With
    Exit Sub

FatalFailure:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    With ThisWorkbook.Worksheets("Task and Data Results")
        .Range("B4").Value2 = "ERROR"
        .Range("B5").Value2 = mCurrentTest & " | " & CStr(capturedNumber) & _
            " | " & capturedSource & " | " & capturedDescription
    End With
End Sub

Private Sub TestHttpSurface()
    Dim badUrlError As Long
    Dim bytes As Variant
    Dim canceledError As Long
    Dim client As ROneCOne
    Dim downloadPath As String
    Dim ensureError As Long
    Dim notFound As ROneCOne
    Dim response As ROneCOne
    Dim source As ROneCOne
    Dim statuses As ROneCOne
    Dim stringFaultError As Long
    Dim task As ROneCOne

    ' Live network contract against https://pokeapi.co/ (authorized test
    ' host). The request bytes move inside WinHTTP while the Task is awaited
    ' cooperatively on Excel's thread.
    mCurrentTest = "TestHttpSurface:get"
    Set client = ROneCOne.HttpClient()
    client.BaseAddress = "https://pokeapi.co/api/v2/"
    client.DefaultRequestHeader "Accept", "application/json"
    AssertEqual "http default timeout", 30000&, client.Timeout
    client.Timeout = 20000
    Set response = client.GetAsync("pokemon/pikachu").Await
    AssertEqual "http get status", 200&, response.StatusCode
    AssertEqual "http get reason", "OK", response.ReasonPhrase
    AssertTrue "http get success flag", response.IsSuccessStatusCode
    AssertTrue "http get body", InStr(1, response.Content, """pikachu""") > 0
    AssertTrue "http content-type header", InStr(1, _
        response.Header("Content-Type"), "application/json") > 0
    AssertTrue "http ensure success returns self", _
        response.EnsureSuccessStatusCode Is response

    mCurrentTest = "TestHttpSurface:getString"
    AssertTrue "http get string", InStr(1, _
        client.GetStringAsync("pokemon/ditto").Await, """ditto""") > 0

    mCurrentTest = "TestHttpSurface:bytes"
    bytes = client.GetByteArrayAsync("pokemon/meowth").Await
    AssertTrue "http byte array", UBound(bytes) > 100

    mCurrentTest = "TestHttpSurface:downloadFile"
    downloadPath = ThisWorkbook.Path & "\ROneCOne_Download.json"
    AssertEqual "http download returns path", downloadPath, _
        client.DownloadFileAsync("pokemon/pikachu", downloadPath).Await
    AssertTrue "http download exists", ROneCOne.File.Exists(downloadPath)
    AssertTrue "http download content", InStr(1, _
        ROneCOne.File.ReadAllText(downloadPath), """pikachu""") > 0
    ROneCOne.File.Delete downloadPath

    mCurrentTest = "TestHttpSurface:whenAll"
    Set statuses = ROneCOne.Task.WhenAll( _
        client.GetAsync("pokemon/bulbasaur"), _
        client.GetAsync("pokemon/charmander"), _
        client.GetAsync("pokemon/squirtle")).Await
    AssertEqual "http whenall count", 3&, statuses.Count
    AssertEqual "http whenall first", 200&, statuses.Item(0).StatusCode
    AssertEqual "http whenall last", 200&, statuses.Item(2).StatusCode

    mCurrentTest = "TestHttpSurface:notFound"
    Set notFound = client.GetAsync( _
        "pokemon/definitely-not-a-pokemon-9999").Await
    AssertEqual "http 404 status", 404&, notFound.StatusCode
    AssertFalse "http 404 success flag", notFound.IsSuccessStatusCode
    ensureError = 0
    On Error Resume Next
    notFound.EnsureSuccessStatusCode
    ensureError = Err.Number
    On Error GoTo 0
    AssertEqual "http ensure success raises", _
        ROneCOne.HttpRequestError, ensureError
    stringFaultError = 0
    On Error Resume Next
    client.GetStringAsync("pokemon/definitely-not-a-pokemon-9999").Await
    stringFaultError = Err.Number
    On Error GoTo 0
    AssertEqual "http get string faults on 404", _
        ROneCOne.HttpRequestError, stringFaultError

    mCurrentTest = "TestHttpSurface:sendVerb"
    Set response = client.SendAsync("POST", "pokemon", _
        "{""probe"":true}", "application/json").Await
    AssertTrue "http post reaches server", response.StatusCode >= 400
    Set response = client.PatchAsync("pokemon/pikachu", _
        "{""probe"":true}", "application/json").Await
    AssertTrue "http patch reaches server", response.StatusCode >= 400
    ' A verb WinHTTP has no name for still transmits; the server answers it.
    Set response = client.SendAsync("FROBNICATE", "pokemon/pikachu").Await
    AssertEqual "http custom verb reaches server", 405&, response.StatusCode

    mCurrentTest = "TestHttpSurface:cancellation"
    Set source = ROneCOne.CancellationTokenSource
    source.Cancel
    Set task = client.GetAsync("pokemon/eevee", source.Token)
    canceledError = 0
    On Error Resume Next
    task.Await
    canceledError = Err.Number
    On Error GoTo 0
    AssertTrue "http canceled task raises", canceledError <> 0
    AssertTrue "http canceled task state", task.IsCanceled

    mCurrentTest = "TestHttpSurface:arguments"
    badUrlError = 0
    On Error Resume Next
    ROneCOne.HttpClient().GetAsync "pokemon/pikachu"
    badUrlError = Err.Number
    On Error GoTo 0
    AssertEqual "http relative url needs base", _
        ROneCOne.InvalidArgumentError, badUrlError
End Sub

Private Sub TestConstraintMaintenance()
    Dim capturedError As Long
    Dim table As ROneCOne

    ' Duplicate detection stays exact while the constraint indexes maintain
    ' themselves incrementally on add and lazily after key edits or deletes.
    Set table = ROneCOne.DataTable("ConstraintChecks")
    table.Column("Id", vbLong).AsPrimaryKey
    table.Column("Code", vbString).AsUnique
    table.Column "Score", vbLong
    table.LoadRow Array(1, "A", 10)
    table.LoadRow Array(2, "B", 20)

    capturedError = 0
    On Error Resume Next
    table.LoadRow Array(1, "C", 30)
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "constraint duplicate key rejected", _
        ROneCOne.InvalidOperationError, capturedError
    AssertEqual "constraint count after key reject", 2&, table.Rows.Count

    capturedError = 0
    On Error Resume Next
    table.LoadRow Array(3, "A", 30)
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "constraint duplicate unique rejected", _
        ROneCOne.InvalidOperationError, capturedError
    AssertEqual "constraint count after unique reject", 2&, table.Rows.Count

    ' An edit that would collide is rejected and the old value restored.
    capturedError = 0
    On Error Resume Next
    table.Rows.Item(1).Item("Code") = "A"
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "constraint duplicate edit rejected", _
        ROneCOne.InvalidOperationError, capturedError
    AssertEqual "constraint edit restored", "B", _
        CStr(table.Rows.Item(1).Item("Code"))

    ' Non-key edits leave lookups untouched; key edits reindex lazily.
    table.Rows.Item(0).Item("Score") = 11
    AssertEqual "constraint find after data edit", 11&, _
        CLng(table.Find(1).Item("Score"))
    table.Rows.Item(1).Item("Id") = 20
    AssertTrue "constraint old key gone", table.Find(2) Is Nothing
    AssertEqual "constraint new key found", "B", _
        CStr(table.Find(20).Item("Code"))

    ' A deleted row leaves both indexes, so its key and unique value free up.
    table.Rows.Item(0).Delete
    AssertTrue "constraint deleted key gone", table.Find(1) Is Nothing
    table.LoadRow Array(1, "A", 99)
    AssertEqual "constraint reuse after delete", 99&, _
        CLng(table.Find(1).Item("Score"))
End Sub

Private Sub TestJsonSurface()
    Dim bound As GenericCustomer
    Dim compact As String
    Dim customers As ROneCOne
    Dim factory As ROneCOne
    Dim jsonErrorNumber As Long
    Dim mapped As ROneCOne
    Dim objects As ROneCOne
    Dim table As ROneCOne
    Dim tree As ROneCOne

    ' Deserialize builds runtime-native values: objects become ordered
    ' String-to-Variant dictionaries, arrays become Variant lists.
    Set tree = ROneCOne.Json.Deserialize( _
        "{""name"":""Ada"",""age"":36,""active"":true,""score"":1.5," & _
        """tags"":[""x"",""y""],""nested"":{""city"":""London""}," & _
        """missing"":null}")
    AssertEqual "json member text", "Ada", tree.Item("name")
    AssertEqual "json member long", 36&, tree.Item("age")
    AssertTrue "json member boolean", CBool(tree.Item("active"))
    AssertEqual "json member double", 1.5, tree.Item("score")
    AssertTrue "json member null", IsNull(tree.Item("missing"))
    AssertEqual "json nested array", "y", tree.Item("tags").Item(1)
    AssertEqual "json nested object", "London", _
        tree.Item("nested").Item("city")

    ' Escapes, backslash-u decoding, and big integers follow the RFC.
    AssertEqual "json escapes", "a""b\c" & ChrW$(233) & vbLf, _
        ROneCOne.Json.Deserialize("""a\""b\\c\u00e9\n""")
    AssertEqual "json long long", 12345678901^, _
        ROneCOne.Json.Deserialize("12345678901")
    jsonErrorNumber = 0
    On Error Resume Next
    ROneCOne.Json.Deserialize("{""a"":01}")
    jsonErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "json leading zero rejected", ROneCOne.JsonError, _
        jsonErrorNumber
    jsonErrorNumber = 0
    On Error Resume Next
    ROneCOne.Json.Deserialize("true false")
    jsonErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "json trailing text rejected", ROneCOne.JsonError, _
        jsonErrorNumber

    ' Serialize round-trips the model with invariant numbers and preserved
    ' member order; pretty output parses back to the identical document.
    compact = "{""a"":1,""b"":[true,null,""x""],""c"":{""d"":2.5}}"
    AssertEqual "json round trip", compact, _
        ROneCOne.Json.Serialize(ROneCOne.Json.Deserialize(compact))
    AssertEqual "json pretty round trip", compact, _
        ROneCOne.Json.Serialize(ROneCOne.Json.Deserialize( _
            ROneCOne.Json.Serialize(ROneCOne.Json.Deserialize(compact), _
                True)))
    AssertEqual "json list sugar", "[1,2,3]", _
        ROneCOne.ListOf(vbLong, 1, 2, 3).ToJson

    ' A JSON array of objects becomes a typed DataTable: dotted columns for
    ' nested objects, JSON text for nested arrays, DBNull for absences.
    Set table = ROneCOne.Json.DeserializeTable( _
        "[{""id"":1,""name"":""Ada"",""meta"":{""city"":""L""}," & _
        """tags"":[1,2]},{""id"":2,""name"":""Bo"",""extra"":true}]")
    AssertEqual "json table rows", 2&, table.Rows.Count
    AssertEqual "json table columns", 5&, table.Columns.Count
    AssertEqual "json table dotted column", "L", _
        CStr(table.Rows.Item(0).Item("meta.city"))
    AssertEqual "json table array cell", "[1,2]", _
        CStr(table.Rows.Item(0).Item("tags"))
    AssertTrue "json table missing cell", _
        IsNull(table.Rows.Item(0).Item("extra"))
    AssertTrue "json table boolean cell", _
        CBool(table.Rows.Item(1).Item("extra"))
    AssertTrue "json table to json", InStr(1, table.ToJson, _
        """meta.city"":""L""") > 0

    ' A nested array path addresses the table inside an envelope.
    AssertEqual "json table at path", 1&, ROneCOne.Json.DeserializeTable( _
        "{""data"":{""items"":[{""v"":7}]}}", "Items", _
        "$.data.items").Rows.Count

    ' Binding: JSON members onto an existing instance, an array of objects
    ' through a factory, and DataTable rows to and from typed objects.
    Set bound = New GenericCustomer
    ROneCOne.Json.DeserializeInto _
        "{""CustomerName"":""Grace"",""Age"":40,""Active"":true}", bound
    AssertEqual "json bind name", "Grace", bound.CustomerName
    AssertEqual "json bind age", 40&, bound.Age
    AssertTrue "json bind boolean", bound.Active

    Set factory = ROneCOne.Func("TestTasksAndData.NewGenericCustomer") _
        .Takes().Returns(vbObject)
    Set objects = ROneCOne.Json.DeserializeObjects( _
        "[{""CustomerName"":""Ada"",""Age"":36}," & _
        "{""CustomerName"":""Bo"",""Age"":50}]", factory)
    AssertEqual "json objects count", 2&, objects.Count
    AssertEqual "json objects typed", "Bo", objects.Item(1).CustomerName
    AssertEqual "json objects bound age", 50&, objects.Item(1).Age

    Set customers = ROneCOne.DataTableFromObjects(objects, _
        Array("CustomerName", "Age"), "Customers")
    AssertEqual "objects to table rows", 2&, customers.Rows.Count
    AssertEqual "objects to table cell", "Ada", _
        CStr(customers.Rows.Item(0).Item("CustomerName"))
    Set mapped = customers.ToObjects(factory)
    AssertEqual "table to objects round trip", 36&, mapped.Item(0).Age
    AssertTrue "table to objects typed list", InStr(1, _
        mapped.GenericTypeName, "GenericCustomer") > 0
End Sub

Public Function NewGenericCustomer() As GenericCustomer
    Set NewGenericCustomer = New GenericCustomer
End Function

Private Sub TestSnapshotCachingAndViewRefresh()
    Dim enumerated As Long
    Dim row As Variant
    Dim Score As Variant
    Dim table As ROneCOne
    Dim view As ROneCOne

    ' Structural snapshot caching: repeated Rows access serves the cached
    ' snapshot, a row add refreshes it, and a field edit stays visible through
    ' the shared row references without invalidating the snapshot.
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:rows"
    Set table = ROneCOne.DataTable("SnapshotCache")
    table.Column "Score", vbLong
    table.LoadRow Array(60&)
    table.LoadRow Array(80&)
    AssertEqual "rows snapshot count", 2&, table.Rows.Count
    table.LoadRow Array(95&)
    AssertEqual "rows snapshot refresh on add", 3&, table.Rows.Count
    table.Rows.Item(0).Item("Score") = 65&
    AssertEqual "field edit via cached snapshot", 65&, _
        CLng(table.Rows.Item(0).Item("Score"))

    ' Schema snapshot caching follows column adds.
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:columns"
    AssertEqual "columns snapshot count", 1&, table.Columns.Count
    table.Column "Note", vbString
    AssertEqual "columns snapshot refresh", 2&, table.Columns.Count
    AssertTrue "late column backfilled", _
        IsNull(table.Rows.Item(0).Item("Note"))

    ' Re-filtering an already-read view must refresh its enumeration and
    ' count; this pinned a live staleness bug in WithFilter and WithSort.
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:viewBuild"
    Set view = ROneCOne.DataView(table) _
        .WithFilter(table.Rows!Score.AtLeast(60&)) _
        .WithSort("Score", True)
    AssertEqual "view count before refilter", 3&, view.Count
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:viewEnum"
    enumerated = 0
    For Each row In view
        enumerated = enumerated + 1
    Next row
    AssertEqual "view enumeration before refilter", 3&, enumerated
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:refilter"
    view.WithFilter table.Rows!Score.AtLeast(80&)
    AssertEqual "view count after refilter", 2&, view.Count
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:refilterEnum"
    enumerated = 0
    For Each row In view
        enumerated = enumerated + 1
    Next row
    AssertEqual "view enumeration after refilter", 2&, enumerated
    AssertEqual "view top row after refilter", 95&, _
        CLng(view.Item(0).Item("Score"))

    ' A row edit that changes filter membership reaches the view through the
    ' data version even though the row set did not change structurally.
    mCurrentTest = "TestSnapshotCachingAndViewRefresh:fieldEdit"
    table.Rows.Item(0).Item("Score") = 99&
    AssertEqual "view count after field edit", 3&, view.Count
End Sub

Private Sub TestExistingRelationValidation()
    Dim child As ROneCOne
    Dim data As ROneCOne
    Dim parent As ROneCOne
    Dim relationError As Long

    Set parent = ROneCOne.DataTable("ExistingParent")
    parent.Column "Id", vbLong
    parent.LoadRow Array(1&)
    Set child = ROneCOne.DataTable("ExistingChild")
    child.Column "ParentId", vbLong
    child.LoadRow Array(99&)
    Set data = ROneCOne.DataSet("ExistingRelations")
    data.AddTable parent
    data.AddTable child

    On Error Resume Next
    data.AddRelation ROneCOne.DataRelation( _
        "ExistingParentChild", parent.Columns("Id"), child.Columns("ParentId"))
    relationError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue "relation rejects existing orphan", relationError <> 0
    AssertEqual "failed relation is not added", 0&, data.Relations.Count
End Sub

Private Sub TestProviderSurface()
    Dim adapter As ROneCOne
    Dim asyncTable As ROneCOne
    Dim command As ROneCOne
    Dim connection As ROneCOne
    Dim connectionString As String
    Dim deleteCommand As ROneCOne
    Dim fillTask As ROneCOne
    Dim insertCommand As ROneCOne
    Dim reader As ROneCOne
    Dim row As ROneCOne
    Dim scalarTask As ROneCOne
    Dim table As ROneCOne
    Dim updateCommand As ROneCOne
    Dim updateTask As ROneCOne
    Dim values As ROneCOne

    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        ThisWorkbook.Path & "\ROneCOne_ProviderFixture.xlsx" & _
        ";Extended Properties=""Excel 12.0 Xml;" & _
        "HDR=YES"";"
    Set connection = ROneCOne.DbConnection(connectionString)
    AssertEqual "provider starts closed", "Closed", connection.State
    AssertEqual "provider async mode", "Native", connection.AsyncMode
    mCurrentTest = "TestProviderSurface.Connect"
    connection.Connect
    AssertEqual "provider opens", "Open", connection.State

    mCurrentTest = "TestProviderSurface.ExecuteReader"
    Set command = ROneCOne.DbCommand( _
        "SELECT [Name], [Score] FROM [Provider Fixture$] " & _
        "ORDER BY [Score] DESC", connection).WithTimeout(30&)
    Set reader = command.ExecuteReader
    AssertEqual "reader field count", 2&, reader.FieldCount
    AssertEqual "reader field name", "Name", reader.GetName(0&)
    AssertEqual "reader field ordinal", 1&, reader.GetOrdinal("Score")
    AssertTrue "reader has rows", reader.HasRows
    AssertTrue "reader first row", reader.Read
    AssertEqual "reader ordered value", "Grace", reader.Item("Name")
    Set values = reader.GetValues
    AssertEqual "reader GetValues", 95#, values.Item(1)
    reader.Disconnect
    AssertTrue "reader closes", reader.IsClosed

    mCurrentTest = "TestProviderSurface.Fill"
    Set adapter = ROneCOne.DbDataAdapter(command)
    Set adapter = adapter.UseTransaction(False).ContinueUpdateOnError(False)
    AssertEqual "adapter starts without update errors", 0&, _
        adapter.LastUpdateErrors.Count
    Set table = ROneCOne.DataTable("ProviderRows")
    AssertEqual "adapter fill count", 2&, adapter.Fill(table)
    AssertEqual "adapter fill rows", 2&, table.Rows.Count
    AssertEqual "filled row state", "Unchanged", table.Rows.Item(0).RowState
    AssertEqual "adapter no changes", 0&, adapter.Update(table)

    mCurrentTest = "TestProviderSurface.FillAsync"
    Set asyncTable = ROneCOne.DataTable("AsyncProviderRows")
    Set fillTask = adapter.FillAsync(asyncTable)
    AssertEqual "adapter FillAsync", 2&, fillTask.Await
    mCurrentTest = "TestProviderSurface.ExecuteScalarAsync"
    Set command = ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM [Provider Fixture$]", connection)
    Set scalarTask = command.ExecuteScalarAsync
    AssertEqual "command ExecuteScalarAsync", 2&, scalarTask.Await

    mCurrentTest = "TestProviderSurface.Update"
    Set updateCommand = ROneCOne.DbCommand( _
        "UPDATE [Provider Fixture$] SET [Score] = ? WHERE [Name] = ?", _
        connection)
    updateCommand.AddParameter ROneCOne.DbParameter( _
        "Score", 0#).FromColumn("Score")
    updateCommand.AddParameter ROneCOne.DbParameter( _
        "Name", vbNullString).FromColumn("Name", "Original")
    Set adapter.UpdateCommand = updateCommand
    table.Rows.Item(0).Item("Score") = 96#
    AssertEqual "adapter updates modified row", 1&, adapter.Update(table)

    mCurrentTest = "TestProviderSurface.UpdateAsync"
    Set insertCommand = ROneCOne.DbCommand( _
        "INSERT INTO [Provider Fixture$] ([Name], [Score]) VALUES (?, ?)", _
        connection)
    insertCommand.AddParameter ROneCOne.DbParameter( _
        "Name", vbNullString).FromColumn("Name")
    insertCommand.AddParameter ROneCOne.DbParameter( _
        "Score", 0#).FromColumn("Score")
    Set adapter.InsertCommand = insertCommand
    Set row = table.LoadRow(Array("Alan", 70#))
    Set updateTask = adapter.UpdateAsync(table)
    AssertEqual "adapter UpdateAsync insert", 1&, updateTask.Await

    mCurrentTest = "TestProviderSurface.Delete"
    Set deleteCommand = ROneCOne.DbCommand( _
        "DELETE FROM [Provider Fixture$] WHERE [Name] = ?", connection)
    deleteCommand.AddParameter ROneCOne.DbParameter( _
        "Name", vbNullString).FromColumn("Name", "Original")
    Set adapter.DeleteCommand = deleteCommand
    AssertTrue "adapter delete command", adapter.DeleteCommand Is deleteCommand
    Set command = ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM [Provider Fixture$]", connection)
    AssertEqual "adapter persisted changes", 3&, command.ExecuteScalar

    mCurrentTest = "TestProviderSurface.Disconnect"
    connection.Disconnect
    AssertEqual "provider closes", "Closed", connection.State
End Sub

Private Sub TestSqlServerProvider()
    Dim canceledError As Long
    Dim command As ROneCOne
    Dim connection As ROneCOne
    Dim elapsed As Double
    Dim failedConnection As ROneCOne
    Dim failureError As Long
    Dim filled As ROneCOne
    Dim insertCount As Long
    Dim reader As ROneCOne
    Dim scalarTask As ROneCOne
    Dim source As ROneCOne
    Dim started As Double
    Dim transaction As ROneCOne
    Dim workCounter As Double

    ' Live SQL Server contract against the local default instance over
    ' MSOLEDBSQL with integrated security, so no credential lives in the
    ' repository. OpenAsync and the execute verbs start natively inside ADO
    ' and the Task polls provider state; the WAITFOR proof below shows the
    ' VBA thread staying free while the server works.
    mCurrentTest = "TestSqlServerProvider.OpenAsync"
    Set connection = ROneCOne.DbConnection(SQL_SERVER_CONNECTION)
    connection.OpenAsync.Await
    AssertEqual "sql open state", "Open", connection.State
    AssertEqual "sql async mode", "Native", connection.AsyncMode

    mCurrentTest = "TestSqlServerProvider.Parameters"
    ROneCOne.DbCommand( _
        "CREATE TABLE #suite_orders (Id int NOT NULL, " & _
        "Customer nvarchar(40) NOT NULL, Total float NOT NULL)", _
        connection).ExecuteNonQuery
    insertCount = ROneCOne.DbCommand( _
        "INSERT INTO #suite_orders VALUES (?, ?, ?), (?, ?, ?)", connection) _
        .WithParameter("Id1", 1&).WithParameter("Customer1", "Ada") _
        .WithParameter("Total1", 12.5) _
        .WithParameter("Id2", 2&).WithParameter("Customer2", "Grace") _
        .WithParameter("Total2", 20#).ExecuteNonQuery
    AssertEqual "sql parameterized insert count", 2&, insertCount
    AssertEqual "sql scalar count", 2&, ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM #suite_orders", connection).ExecuteScalar

    mCurrentTest = "TestSqlServerProvider.ReaderAsync"
    Set reader = ROneCOne.DbCommand( _
        "SELECT Id, Customer, Total FROM #suite_orders ORDER BY Total DESC", _
        connection).ExecuteReaderAsync.Await
    AssertTrue "sql reader async first row", reader.Read
    AssertEqual "sql reader async value", "Grace", reader.Item("Customer")
    AssertEqual "sql reader async double", 20#, reader.Item("Total")
    reader.Disconnect

    mCurrentTest = "TestSqlServerProvider.FillAsync"
    Set filled = ROneCOne.DataTable("SqlOrders")
    AssertEqual "sql fill async count", 2&, ROneCOne.DbDataAdapter( _
        ROneCOne.DbCommand( _
        "SELECT Id, Customer, Total FROM #suite_orders ORDER BY Id", _
        connection)).FillAsync(filled).Await
    AssertEqual "sql filled first customer", "Ada", _
        CStr(filled.Rows.Item(0).Item("Customer"))

    mCurrentTest = "TestSqlServerProvider.Transactions"
    Set transaction = connection.BeginTransaction
    ROneCOne.DbCommand( _
        "UPDATE #suite_orders SET Total = 99 WHERE Id = 1", connection) _
        .ExecuteNonQuery
    transaction.Rollback
    AssertEqual "sql rollback restores", 12.5, ROneCOne.DbCommand( _
        "SELECT Total FROM #suite_orders WHERE Id = 1", connection) _
        .ExecuteScalar
    Set transaction = connection.BeginTransaction
    ROneCOne.DbCommand( _
        "UPDATE #suite_orders SET Total = 30 WHERE Id = 1", connection) _
        .ExecuteNonQuery
    transaction.Commit
    AssertEqual "sql commit persists", 30#, ROneCOne.DbCommand( _
        "SELECT Total FROM #suite_orders WHERE Id = 1", connection) _
        .ExecuteScalar

    mCurrentTest = "TestSqlServerProvider.NativeOverlap"
    Set command = ROneCOne.DbCommand( _
        "SET NOCOUNT ON; WAITFOR DELAY '00:00:00.300'; " & _
        "SELECT COUNT(*) FROM #suite_orders", connection)
    started = Timer
    Set scalarTask = command.ExecuteScalarAsync
    AssertFalse "sql async starts pending", scalarTask.IsCompleted
    ' Busy local work while the server sits in WAITFOR: with native async the
    ' total stays near the 0.3s delay; a blocking execution would serialize
    ' to at least 0.55s and fail the ceiling below.
    workCounter = 0
    Do While ElapsedSecondsSince(started) < 0.25
        workCounter = workCounter + 1
    Loop
    AssertEqual "sql async awaited result", 2&, scalarTask.Await
    elapsed = ElapsedSecondsSince(started)
    AssertTrue "sql async overlapped provider wait", elapsed < 0.45
    AssertTrue "sql busy loop actually ran", workCounter > 1000

    mCurrentTest = "TestSqlServerProvider.Cancellation"
    Set source = ROneCOne.CancellationTokenSource
    Set scalarTask = ROneCOne.DbCommand( _
        "SET NOCOUNT ON; WAITFOR DELAY '00:00:05'; SELECT 1", connection) _
        .ExecuteScalarAsync(source.Token)
    source.Cancel
    canceledError = 0
    On Error Resume Next
    scalarTask.Await
    canceledError = Err.Number
    On Error GoTo 0
    AssertTrue "sql canceled await raises", canceledError <> 0
    AssertTrue "sql canceled task state", scalarTask.IsCanceled
    AssertEqual "sql connection usable after cancel", 2&, ROneCOne.DbCommand( _
        "SELECT COUNT(*) FROM #suite_orders", connection).ExecuteScalar

    mCurrentTest = "TestSqlServerProvider.FailureShapes"
    failureError = 0
    On Error Resume Next
    ROneCOne.DbCommand("SELECT missing_column FROM #suite_orders", _
        connection).ExecuteScalarAsync.Await
    failureError = Err.Number
    On Error GoTo 0
    AssertTrue "sql async execute failure surfaces", failureError <> 0
    Set failedConnection = ROneCOne.DbConnection( _
        "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=tempdb;" & _
        "User ID=ronecone_missing_login;Password=not_a_secret;")
    failureError = 0
    On Error Resume Next
    failedConnection.OpenAsync.Await
    failureError = Err.Number
    On Error GoTo 0
    AssertTrue "sql async open failure surfaces", failureError <> 0
    AssertEqual "sql failed connection state", "Closed", failedConnection.State

    mCurrentTest = "TestSqlServerProvider.Disconnect"
    connection.Disconnect
    AssertEqual "sql connection closes", "Closed", connection.State
End Sub

Private Sub TestTaskLifecycle()
    Dim canceledError As Long
    Dim ignored As Variant
    Dim source As ROneCOne
    Dim taskValue As ROneCOne
    Dim work As ROneCOne

    Set work = ROneCOne.Value(CLng(42)).AsFunc
    Set taskValue = ROneCOne.Task.Run(work)

    AssertFalse "task starts incomplete", taskValue.IsCompleted
    AssertEqual "task await result", CLng(42), taskValue.Await
    AssertTrue "task completes", taskValue.IsCompleted
    AssertEqual "task status", "RanToCompletion", taskValue.Status
    AssertEqual "task result", CLng(42), taskValue.Result

    Set source = ROneCOne.CancellationTokenSource
    source.Cancel
    Set taskValue = ROneCOne.Task.Run(work, source.Token)
    On Error Resume Next
    ignored = taskValue.Await
    canceledError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "pre-canceled cooperative Await", _
        ROneCOne.OperationCanceledError, canceledError
    AssertTrue "pre-canceled cooperative state", taskValue.IsCanceled
End Sub

Private Sub TestTaskCombinators()
    Dim allTask As ROneCOne
    Dim completed As ROneCOne
    Dim completion As ROneCOne
    Dim firstTask As ROneCOne
    Dim pending As ROneCOne
    Dim results As ROneCOne
    Dim secondTask As ROneCOne
    Dim tasks As ROneCOne
    Dim winner As ROneCOne
    Dim faulted As ROneCOne
    Dim faultNumber As Long
    Dim ignored As Variant

    Set firstTask = ROneCOne.Task.FromResult(CLng(10))
    Set secondTask = ROneCOne.Task.FromResult(CLng(20))
    Set tasks = ROneCOne.ListOf(ROneCOne.Task, firstTask, secondTask)
    Set allTask = ROneCOne.Task.WhenAll(tasks)
    Set results = allTask.Await

    AssertEqual "WhenAll result count", CLng(2), results.Count
    AssertEqual "WhenAll preserves order", CLng(20), results.Item(1)
    AssertTrue "WhenAny completes", ROneCOne.Task.WhenAny(tasks).Await Is firstTask

    Set results = ROneCOne.Task.WhenAll(firstTask, secondTask).Await
    AssertEqual "WhenAll accepts direct Tasks", 2&, results.Count
    Set winner = ROneCOne.Task.WhenAny(secondTask, firstTask).Await
    AssertTrue "WhenAny accepts direct Tasks", winner Is secondTask

    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    Set pending = completion.Task
    Set completed = ROneCOne.Task.FromResult(42&)
    Set tasks = ROneCOne.ListOf(ROneCOne.Task, pending, completed)
    Set winner = ROneCOne.Task.WhenAny(tasks).Await
    AssertTrue "WhenAny selects a completed task", winner Is completed
    AssertEqual "WhenAny completed result", 42&, winner.Result

    Set firstTask = ROneCOne.Value(1&).Divide(0&).AsFunc
    Set firstTask = firstTask.Returns(vbLong)
    Set secondTask = ROneCOne.Value(2&).Divide(0&).AsFunc
    Set secondTask = secondTask.Returns(vbLong)
    Set faulted = ROneCOne.Task.WhenAll( _
        ROneCOne.Task.Run(firstTask), _
        ROneCOne.Task.Run(secondTask))
    On Error Resume Next
    ignored = faulted.Await
    faultNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue "WhenAll await reports a fault", faultNumber <> 0
    AssertEqual "WhenAll exposes AggregateException", _
        "AggregateException", faulted.Exception.ExceptionType
    AssertEqual "WhenAll retains every fault", 2&, _
        faulted.Exception.InnerExceptions.Count
End Sub

Private Sub TestCancellation()
    Dim callback As ROneCOne
    Dim registration As ROneCOne
    Dim secondRegistration As ROneCOne
    Dim source As ROneCOne
    Dim token As ROneCOne
    Dim usingResult As Variant

    Set source = ROneCOne.CancellationTokenSource
    Set token = source.Token
    AssertFalse "token starts active", token.IsCancellationRequested
    source.Cancel
    AssertTrue "token observes cancellation", token.IsCancellationRequested

    DelegateProcedures.ResetTrace
    Set source = ROneCOne.CancellationTokenSource
    Set registration = source.Token.Register(ROneCOne.Action( _
        "DelegateProcedures.RecordCancellation").Takes)
    registration.Dispose
    source.Cancel
    AssertEqual "disposed cancellation registration", vbNullString, _
        DelegateProcedures.CurrentTrace

    DelegateProcedures.ResetTrace
    Set source = ROneCOne.CancellationTokenSource
    Set callback = ROneCOne.Action( _
        "DelegateProcedures.RecordCancellation").Takes
    Set registration = source.Token.Register(callback)
    Set secondRegistration = source.Token.Register(callback)
    registration.Dispose
    source.Cancel
    AssertEqual "duplicate registration removes one callback", "canceled|", _
        DelegateProcedures.CurrentTrace
    secondRegistration.Dispose

    DelegateProcedures.ResetTrace
    Set source = ROneCOne.CancellationTokenSource
    Set registration = source.Token.Register(ROneCOne.Action( _
        "DelegateProcedures.RecordCancellation").Takes)
    usingResult = ROneCOne.Using(registration).Run( _
        ROneCOne.Value(42&).AsFunc)
    source.Cancel
    AssertEqual "Using returns Func result", 42&, usingResult
    AssertEqual "Using disposes the resource", vbNullString, _
        DelegateProcedures.CurrentTrace
End Sub

Private Sub TestTaskCompletionAndProgress()
    Dim completion As ROneCOne
    Dim progress As ROneCOne
    Dim taskValue As ROneCOne

    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    Set taskValue = completion.Task
    AssertFalse "completion source starts pending", taskValue.IsCompleted
    AssertTrue "completion source TrySetResult", completion.TrySetResult(42&)
    AssertFalse "completion source rejects second result", _
        completion.TrySetResult(99&)
    AssertEqual "completion source result", 42&, taskValue.Await

    DelegateProcedures.ResetTotal
    Set progress = ROneCOne.ProgressOf( _
        vbLong, ROneCOne.Action( _
            "DelegateProcedures.AccumulateLong").Takes(vbLong))
    progress.Report 7&
    AssertEqual "typed progress", 7&, DelegateProcedures.CurrentTotal

    AssertTrue "Task.CompletedTask", ROneCOne.Task.CompletedTask.IsCompleted
    AssertEqual "Task.FromResult", 9&, ROneCOne.Task.FromResult(9&).Await
End Sub

Private Sub TestTaskTimeoutAndContinuation()
    Dim continuation As ROneCOne
    Dim delayed As ROneCOne
    Dim ignored As Variant
    Dim resultTask As ROneCOne
    Dim source As ROneCOne
    Dim timeoutError As Long
    Dim waitTask As ROneCOne

    Set delayed = ROneCOne.Task.Delay(20&)
    AssertTrue "pending Task.Exception is Nothing", delayed.Exception Is Nothing
    AssertFalse "delay timeout", delayed.Wait(1&)
    AssertTrue "delay eventually completes", delayed.Wait(100&)

    ignored = ROneCOne.Task.YieldOnce.Await
    AssertTrue "Task.YieldOnce completes", True

    Set waitTask = ROneCOne.Task.Delay(20&).WaitAsync(100&)
    ignored = waitTask.Await
    AssertTrue "WaitAsync completes before timeout", waitTask.IsCompleted

    Set waitTask = ROneCOne.Task.Delay(50&).WaitAsync(1&)
    On Error Resume Next
    ignored = waitTask.Await
    timeoutError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "WaitAsync raises TimeoutException", _
        ROneCOne.TimeoutError, timeoutError
    AssertEqual "WaitAsync exception type", "TimeoutException", _
        waitTask.Exception.InnerExceptions.Item(0).ExceptionType

    Set continuation = ROneCOne.Func( _
        "DelegateProcedures.TaskResultPlusOne") _
        .Takes(ROneCOne.Task) _
        .Returns(vbLong)
    Set resultTask = ROneCOne.Task.FromResult(10&).ContinueWith(continuation)
    AssertEqual "ContinueWith", 11&, resultTask.Await

    DelegateProcedures.ResetTrace
    Set source = ROneCOne.CancellationTokenSource
    source.Token.Register ROneCOne.Action( _
        "DelegateProcedures.RecordCancellation").Takes
    source.Cancel
    AssertEqual "cancellation callback", "canceled|", _
        DelegateProcedures.CurrentTrace
End Sub

Private Sub TestDataTable()
    Dim duplicateError As Long
    Dim found As ROneCOne
    Dim row As ROneCOne
    Dim table As ROneCOne

    mCurrentTest = "TestDataTable.Table"
    Set table = ROneCOne.DataTable("Customers")
    mCurrentTest = "TestDataTable.Columns"
    table.AddColumn ROneCOne.DataColumn("Id", vbLong).AsUnique
    table.AddColumn ROneCOne.DataColumn("Name", vbString)
    mCurrentTest = "TestDataTable.PrimaryKey"
    Set table.PrimaryKey = table.Columns.Where("ColumnName").EqualTo("Id").ToList

    mCurrentTest = "TestDataTable.NewRow"
    Set row = table.NewRow
    row.Item("Id") = CLng(1)
    row.Item("Name") = "Ada"
    mCurrentTest = "TestDataTable.AddRow"
    table.AddRow row

    AssertEqual "table column count", CLng(2), table.Columns.Count
    AssertEqual "table row count", CLng(1), table.Rows.Count
    AssertEqual "row field by name", "Ada", table.Rows.Item(0).Item("Name")
    mCurrentTest = "TestDataTable.Find.Call"
    Set found = table.Find(CLng(1))
    mCurrentTest = "TestDataTable.Find.Return"
    AssertTrue "primary-key row returned", Not found Is Nothing
    AssertTrue "primary-key identity", found Is row
    mCurrentTest = "TestDataTable.Find.Item"
    If found Is row Then
        AssertEqual "primary-key find", "Ada", found.Item("Name")
    End If

    mCurrentTest = "TestDataTable.AcceptChanges"
    table.AcceptChanges
    AssertEqual "accepted row state", "Unchanged", row.RowState
    mCurrentTest = "TestDataTable.Modify"
    row.Item("Name") = "Augusta"
    AssertEqual "modified row state", "Modified", row.RowState
    mCurrentTest = "TestDataTable.RejectChanges"
    row.RejectChanges
    AssertEqual "reject restores original", "Ada", row.Item("Name")

    mCurrentTest = "TestDataTable.Duplicate"
    On Error Resume Next
    Set row = table.NewRow
    row.Item("Id") = CLng(1)
    row.Item("Name") = "Duplicate"
    table.AddRow row
    duplicateError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue "unique column rejects duplicate", duplicateError <> 0
End Sub

Private Sub TestDataSetRelations()
    Dim child As ROneCOne
    Dim children As ROneCOne
    Dim customers As ROneCOne
    Dim data As ROneCOne
    Dim orders As ROneCOne
    Dim parent As ROneCOne
    Dim relation As ROneCOne

    Set customers = ROneCOne.DataTable("Customers")
    customers.AddColumn ROneCOne.DataColumn("Id", vbLong).AsUnique
    Set parent = customers.NewRow
    parent.Item("Id") = CLng(1)
    customers.AddRow parent

    Set orders = ROneCOne.DataTable("Orders")
    orders.AddColumn ROneCOne.DataColumn("CustomerId", vbLong)
    Set child = orders.NewRow
    child.Item("CustomerId") = CLng(1)
    orders.AddRow child

    Set data = ROneCOne.DataSet("Sales")
    data.AddTable customers
    data.AddTable orders
    mCurrentTest = "TestDataSetRelations.CreateRelation"
    Set relation = ROneCOne.DataRelation( _
        "CustomerOrders", customers.Columns.Item("Id"), _
        orders.Columns.Item("CustomerId"))
    mCurrentTest = "TestDataSetRelations.AddRelation"
    data.AddRelation relation

    mCurrentTest = "TestDataSetRelations.GetChildren"
    Set children = parent.GetChildRows("CustomerOrders")
    AssertEqual "relation child navigation", CLng(1), children.Count
    mCurrentTest = "TestDataSetRelations.GetParent"
    AssertTrue "relation parent navigation", _
        child.GetParentRow("CustomerOrders") Is parent
    AssertEqual "dataset table count", CLng(2), data.Tables.Count
End Sub

Private Sub TestAdvancedDataTable()
    Dim changes As ROneCOne
    Dim composite As ROneCOne
    Dim row As ROneCOne
    Dim selected As ROneCOne
    Dim table As ROneCOne

    Set table = ROneCOne.DataTable("People")
    mCurrentTest = "TestAdvancedDataTable.IdSchema"
    table.Column("Id", vbLong).AutoNumber(100&, 10&).AsPrimaryKey
    mCurrentTest = "TestAdvancedDataTable.NameSchema"
    table.Column("Name", vbString).WithDefault "Unknown"
    mCurrentTest = "TestAdvancedDataTable.NoteSchema"
    table.Column "Note", vbString
    mCurrentTest = "TestAdvancedDataTable.RowAdd"
    Set row = table.Row("Ada", ROneCOne.DBNull).Add

    AssertEqual "auto increment", 100&, row.Item("Id")
    AssertEqual "fluent row value", "Ada", row.Item("Name")
    AssertTrue "explicit DBNull", IsNull(row.Item("Note"))
    AssertTrue "indexed primary-key Find", table.Find(100&) Is row
    Set selected = table.SelectRows(table.Rows!Name.EqualTo("Ada"))
    AssertEqual "table SelectRows", 1&, selected.Count

    table.AcceptChanges
    row.Item("Name") = "Ada"
    Set changes = table.GetChanges
    AssertEqual "table GetChanges", 1&, changes.Rows.Count
    row.Delete
    AssertEqual "row Delete", "Deleted", row.RowState

    Set composite = ROneCOne.DataTable("Composite")
    mCurrentTest = "TestAdvancedDataTable.CompositeSchema"
    composite.Column "Region", vbString
    composite.Column "Id", vbLong
    Set composite.PrimaryKey = ROneCOne.ListFrom( _
        composite.Columns("Region"), composite.Columns("Id"))
    mCurrentTest = "TestAdvancedDataTable.CompositeRow"
    Set row = composite.Row("US", 7&).Add
    AssertTrue "composite primary-key Find", _
        composite.Find("US", 7&) Is row
End Sub

Private Sub TestCompositeRelations()
    Dim child As ROneCOne
    Dim childRow As ROneCOne
    Dim data As ROneCOne
    Dim parent As ROneCOne
    Dim parentRow As ROneCOne

    Set parent = ROneCOne.DataTable("CompositeParent")
    parent.Column "Region", vbString
    parent.Column "Id", vbLong
    Set child = ROneCOne.DataTable("CompositeChild")
    child.Column "Region", vbString
    child.Column "ParentId", vbLong
    child.Column "Name", vbString
    Set data = ROneCOne.DataSet("CompositeRelations")
    data.AddTable parent
    data.AddTable child
    data.AddRelation ROneCOne.DataRelation( _
        "CompositeParentChild", _
        ROneCOne.ListFrom(parent.Columns("Region"), parent.Columns("Id")), _
        ROneCOne.ListFrom(child.Columns("Region"), child.Columns("ParentId")))
    Set parentRow = parent.Row("US", 1&).Add
    Set childRow = child.Row("US", 1&, "Order").Add
    AssertEqual "composite relation child navigation", 1&, _
        parentRow.GetChildRows("CompositeParentChild").Count
    AssertTrue "composite relation parent navigation", _
        childRow.GetParentRow("CompositeParentChild") Is parentRow
End Sub

Private Sub TestDataViewAndMerge()
    Dim copied As ROneCOne
    Dim merged As ROneCOne
    Dim row As ROneCOne
    Dim Score As Variant
    Dim table As ROneCOne
    Dim view As ROneCOne

    Set table = ROneCOne.DataTable("Scores")
    table.Column "Name", vbString
    table.Column "Score", vbLong
    Set row = table.LoadRow(Array("Ada", 90&))
    Set row = table.LoadRow(Array("Grace", 95&))
    Set row = table.LoadRow(Array("Alan", 70&))

    Set view = ROneCOne.DataView(table) _
        .WithFilter(table.Rows!Score.AtLeast(80&)) _
        .WithSort("Score", True)
    AssertEqual "DataView filter", 2&, view.Count
    AssertEqual "DataView sort", "Grace", view.Item(0).Item("Name")

    Set copied = table.Copy
    AssertEqual "DataTable Copy", 3&, copied.Rows.Count
    Set merged = table.CloneTable
    merged.Merge copied
    AssertEqual "DataTable Merge", 3&, merged.Rows.Count
End Sub

Private Sub TestRangeBridge()
    Dim numbers As ROneCOne
    Dim readBack As ROneCOne
    Dim source As ROneCOne
    Dim sheet As Object
    Dim table As ROneCOne
    Dim written As Long

    Set sheet = ThisWorkbook.Worksheets("Task and Data Results")
    sheet.Range("H1:Q100").ClearContents

    ' Seed a small headered grid the way a user's worksheet would hold it.
    sheet.Range("H1:J1").Value = Array("Name", "Score", "Active")
    sheet.Range("H2:J2").Value = Array("Ada", 90&, True)
    sheet.Range("H3:J3").Value = Array("Grace", 95&, False)

    ' Read the whole block into a DataTable in one bulk call.
    Set table = ROneCOne.DataTableFromRange(sheet.Range("H1:J3"))
    AssertEqual "range table columns", 3&, table.Columns.Count
    AssertEqual "range table rows", 2&, table.Rows.Count
    AssertEqual "range table text cell", "Grace", table.Rows.Item(1).Item("Name")
    AssertEqual "range table number cell", 90&, table.Rows.Item(0).Item("Score")

    ' Write the table back out and read it in again: a full round trip.
    written = table.ToRange(sheet.Range("L1"))
    AssertEqual "table ToRange row count", 2&, written
    Set readBack = ROneCOne.DataTableFromRange(sheet.Range("L1:N3"))
    AssertEqual "round-trip rows", 2&, readBack.Rows.Count
    AssertEqual "round-trip cell", "Ada", readBack.Rows.Item(0).Item("Name")

    ' A single column reads into a List.
    Set source = ROneCOne.ListFromRange(sheet.Range("I2:I3"))
    AssertEqual "list from range count", 2&, source.Count
    AssertEqual "list from range value", 90&, source.Item(0)

    ' A scalar sequence writes itself down a column.
    Set numbers = ROneCOne.ListOf(vbLong, 3, 1, 2)
    written = numbers.Order.ToList.ToRange(sheet.Range("P1"))
    AssertEqual "sequence ToRange count", 3&, written
    AssertEqual "sequence ToRange first cell", 1&, sheet.Range("P1").Value

    sheet.Range("H1:Q100").ClearContents
End Sub

Private Sub TestRelationConstraints()
    Dim child As ROneCOne
    Dim childRow As ROneCOne
    Dim data As ROneCOne
    Dim parent As ROneCOne
    Dim parentRow As ROneCOne
    Dim relationError As Long

    Set parent = ROneCOne.DataTable("Parent")
    parent.Column("Id", vbLong).AsUnique
    Set child = ROneCOne.DataTable("Child")
    child.Column "ParentId", vbLong
    Set data = ROneCOne.DataSet("Relations")
    data.AddTable parent
    data.AddTable child
    data.AddRelation ROneCOne.DataRelation( _
        "ParentChild", parent.Columns("Id"), child.Columns("ParentId"))

    Set parentRow = parent.LoadRow(Array(1&))
    Set childRow = child.LoadRow(Array(1&))
    AssertEqual "foreign key valid", 1&, child.Rows.Count

    On Error Resume Next
    Set childRow = child.LoadRow(Array(99&))
    relationError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue "foreign key rejects orphan", relationError <> 0
End Sub

Private Sub TestFileSystemSurface()
    Dim bytes As Variant
    Dim capturedError As Long
    Dim directories As ROneCOne
    Dim files As ROneCOne
    Dim lines As ROneCOne
    Dim payload(0 To 2) As Byte
    Dim sample As String
    Dim testRoot As String

    ' Every path lives under one scratch root next to the workbook, created
    ' first and removed last, so the suite leaves nothing behind.
    testRoot = ThisWorkbook.Path & "\ROneCOne_FileTests"
    If ROneCOne.Directory.Exists(testRoot) Then
        ROneCOne.Directory.Delete testRoot, True
    End If
    ROneCOne.Directory.CreateDirectory testRoot & "\inner\deep"
    AssertTrue "fs nested create", _
        ROneCOne.Directory.Exists(testRoot & "\inner\deep")

    mCurrentTest = "TestFileSystemSurface.Text"
    sample = "he" & ChrW$(233) & "llo " & ChrW$(8364)
    ROneCOne.File.WriteAllText testRoot & "\sample.txt", sample
    AssertEqual "fs utf8 round trip", sample, _
        ROneCOne.File.ReadAllText(testRoot & "\sample.txt")
    bytes = ROneCOne.File.ReadAllBytes(testRoot & "\sample.txt")
    AssertEqual "fs utf8 write has no bom", 104&, CLng(bytes(0))
    ROneCOne.File.AppendAllText testRoot & "\sample.txt", "!"
    AssertEqual "fs append", sample & "!", _
        ROneCOne.File.ReadAllText(testRoot & "\sample.txt")

    mCurrentTest = "TestFileSystemSurface.Encodings"
    ROneCOne.File.WriteAllText testRoot & "\sample16.txt", sample, "utf-16"
    bytes = ROneCOne.File.ReadAllBytes(testRoot & "\sample16.txt")
    AssertEqual "fs utf16 write keeps bom", 255&, CLng(bytes(0))
    AssertEqual "fs bom decides decoder", sample, _
        ROneCOne.File.ReadAllText(testRoot & "\sample16.txt")
    ROneCOne.File.WriteAllText testRoot & "\ansi.txt", _
        "h" & ChrW$(233), "windows-1252"
    bytes = ROneCOne.File.ReadAllBytes(testRoot & "\ansi.txt")
    AssertEqual "fs ansi single byte", 233&, CLng(bytes(1))
    AssertEqual "fs ansi round trip", "h" & ChrW$(233), _
        ROneCOne.File.ReadAllText(testRoot & "\ansi.txt", "windows-1252")
    capturedError = 0
    On Error Resume Next
    ROneCOne.File.ReadAllText testRoot & "\sample.txt", "latin-5"
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "fs unknown encoding rejected", _
        ROneCOne.InvalidArgumentError, capturedError

    mCurrentTest = "TestFileSystemSurface.Lines"
    ROneCOne.File.WriteAllLines testRoot & "\lines.txt", _
        Array("alpha", "beta", "gamma")
    Set lines = ROneCOne.File.ReadAllLines(testRoot & "\lines.txt")
    AssertEqual "fs lines count", 3&, lines.Count
    AssertEqual "fs lines first", "alpha", lines.Item(0)
    AssertEqual "fs lines last", "gamma", lines.Item(2)
    AssertTrue "fs lines terminate", _
        Right$(ROneCOne.File.ReadAllText(testRoot & "\lines.txt"), 2) = vbCrLf

    mCurrentTest = "TestFileSystemSurface.Bytes"
    payload(0) = 1
    payload(1) = 2
    payload(2) = 3
    ROneCOne.File.WriteAllBytes testRoot & "\payload.bin", payload
    bytes = ROneCOne.File.ReadAllBytes(testRoot & "\payload.bin")
    AssertEqual "fs bytes bound", 2&, CLng(UBound(bytes))
    AssertEqual "fs bytes value", 3&, CLng(bytes(2))

    mCurrentTest = "TestFileSystemSurface.CopyMoveDelete"
    ROneCOne.File.Copy testRoot & "\sample.txt", testRoot & "\copy.txt"
    AssertTrue "fs copy exists", ROneCOne.File.Exists(testRoot & "\copy.txt")
    capturedError = 0
    On Error Resume Next
    ROneCOne.File.Copy testRoot & "\sample.txt", testRoot & "\copy.txt"
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "fs copy refuses overwrite", ROneCOne.IOError, capturedError
    ROneCOne.File.Copy testRoot & "\sample.txt", testRoot & "\copy.txt", True
    ROneCOne.File.Move testRoot & "\copy.txt", testRoot & "\moved.txt"
    AssertTrue "fs move target", ROneCOne.File.Exists(testRoot & "\moved.txt")
    AssertFalse "fs move source gone", _
        ROneCOne.File.Exists(testRoot & "\copy.txt")
    ROneCOne.File.Delete testRoot & "\moved.txt"
    AssertFalse "fs delete", ROneCOne.File.Exists(testRoot & "\moved.txt")
    capturedError = 0
    On Error Resume Next
    ROneCOne.File.Delete testRoot & "\moved.txt"
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "fs delete missing is silent", 0&, capturedError
    capturedError = 0
    On Error Resume Next
    ROneCOne.File.ReadAllText testRoot & "\moved.txt"
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "fs read missing raises", ROneCOne.IOError, capturedError

    mCurrentTest = "TestFileSystemSurface.Enumeration"
    ROneCOne.File.WriteAllText testRoot & "\inner\note.txt", "inner note"
    Set files = ROneCOne.Directory.GetFiles(testRoot, "*.txt")
    AssertEqual "fs top files", 4&, files.Count
    AssertEqual "fs files sorted", "ansi.txt", _
        ROneCOne.Path.GetFileName(CStr(files.Item(0)))
    Set files = ROneCOne.Directory.GetFiles(testRoot, "*.txt", True)
    AssertEqual "fs recursive files", 5&, files.Count
    Set files = ROneCOne.Directory.GetFiles(testRoot, "note.*", True)
    AssertEqual "fs pattern match", 1&, files.Count
    Set directories = ROneCOne.Directory.GetDirectories(testRoot)
    AssertEqual "fs top directories", 1&, directories.Count
    Set directories = ROneCOne.Directory.GetDirectories(testRoot, "*", True)
    AssertEqual "fs recursive directories", 2&, directories.Count

    mCurrentTest = "TestFileSystemSurface.Paths"
    AssertEqual "path combine", "C:\data\in\file.txt", _
        ROneCOne.Path.Combine("C:\data", "in", "file.txt")
    AssertEqual "path combine rooted reset", "D:\other", _
        ROneCOne.Path.Combine("C:\data", "D:\other")
    AssertEqual "path combine keeps separator", "C:\data\x.txt", _
        ROneCOne.Path.Combine("C:\data\", "x.txt")
    AssertEqual "path file name", "b.txt", _
        ROneCOne.Path.GetFileName("C:\a\b.txt")
    AssertEqual "path directory name", "C:\a", _
        ROneCOne.Path.GetDirectoryName("C:\a\b.txt")
    AssertEqual "path extension", ".txt", _
        ROneCOne.Path.GetExtension("C:\a\b.txt")
    AssertEqual "path no extension", vbNullString, _
        ROneCOne.Path.GetExtension("C:\a\b")
    AssertEqual "path name without extension", "b", _
        ROneCOne.Path.GetFileNameWithoutExtension("C:\a\b.txt")
    AssertEqual "path dot file", vbNullString, _
        ROneCOne.Path.GetFileNameWithoutExtension(".gitignore")
    AssertEqual "path change extension", "C:\a\b.json", _
        ROneCOne.Path.ChangeExtension("C:\a\b.txt", "json")
    AssertEqual "path remove extension", "C:\a\b", _
        ROneCOne.Path.ChangeExtension("C:\a\b.txt", vbNullString)
    AssertTrue "path temp trailing separator", _
        Right$(ROneCOne.Path.GetTempPath(), 1) = "\"
    AssertTrue "path full path is rooted", _
        Mid$(ROneCOne.Path.GetFullPath("sample.txt"), 2, 1) = ":"

    mCurrentTest = "TestFileSystemSurface.Cleanup"
    capturedError = 0
    On Error Resume Next
    ROneCOne.Directory.Delete testRoot
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "fs non-recursive delete refuses", _
        ROneCOne.IOError, capturedError
    ROneCOne.Directory.Delete testRoot, True
    AssertFalse "fs root removed", ROneCOne.Directory.Exists(testRoot)
End Sub

Private Sub TestCsvSurface()
    Dim csvErrorNumber As Long
    Dim document As String
    Dim filePath As String
    Dim roundTripped As ROneCOne
    Dim Score As Variant
    Dim table As ROneCOne
    Dim view As ROneCOne

    ' The writer quotes only where RFC 4180 demands it, keeps numbers and
    ' dates locale-safe, and distinguishes a database null (empty field)
    ' from an empty string (quoted pair).
    mCurrentTest = "TestCsvSurface.Serialize"
    Set table = ROneCOne.DataTable("Orders")
    table.Column "Id", vbLong
    table.Column "Customer", vbString
    table.Column "Total", vbDouble
    table.Column "Active", vbBoolean
    table.Column "Placed", vbDate
    table.Column "Note", vbString
    table.LoadRow Array(1, "Ada, Ltd", 12.5, True, _
        DateSerial(2026, 7, 23) + TimeSerial(10, 30, 0), ROneCOne.DBNull)
    table.LoadRow Array(2, "Grace ""G""", 20#, False, _
        DateSerial(2026, 1, 2), vbNullString)
    document = table.ToCsv
    AssertEqual "csv document", _
        "Id,Customer,Total,Active,Placed,Note" & vbCrLf & _
        "1,""Ada, Ltd"",12.5,true,2026-07-23T10:30:00," & vbCrLf & _
        "2,""Grace """"G"""""",20,false,2026-01-02T00:00:00,""""" & vbCrLf, _
        document

    mCurrentTest = "TestCsvSurface.RoundTrip"
    Set roundTripped = ROneCOne.Csv.DeserializeTable(document, "Orders")
    AssertEqual "csv round trip rows", 2&, roundTripped.Rows.Count
    AssertEqual "csv long cell", 1&, roundTripped.Rows.Item(0).Item("Id")
    AssertEqual "csv comma text", "Ada, Ltd", _
        CStr(roundTripped.Rows.Item(0).Item("Customer"))
    AssertEqual "csv double cell", 12.5, _
        roundTripped.Rows.Item(0).Item("Total")
    AssertEqual "csv boolean cell", True, _
        roundTripped.Rows.Item(0).Item("Active")
    AssertEqual "csv date cell", _
        DateSerial(2026, 7, 23) + TimeSerial(10, 30, 0), _
        roundTripped.Rows.Item(0).Item("Placed")
    AssertTrue "csv null survives", _
        IsNull(roundTripped.Rows.Item(0).Item("Note"))
    AssertEqual "csv empty string survives", vbNullString, _
        CStr(roundTripped.Rows.Item(1).Item("Note"))
    AssertEqual "csv quote escape", "Grace ""G""", _
        CStr(roundTripped.Rows.Item(1).Item("Customer"))

    mCurrentTest = "TestCsvSurface.Inference"
    Set table = ROneCOne.Csv.DeserializeTable( _
        "A,B,C,D" & vbLf & _
        "001,5,true,2026-02-31" & vbLf & _
        "abc,9999999999,false,note")
    AssertEqual "csv leading zero stays text", "001", _
        CStr(table.Rows.Item(0).Item("A"))
    AssertEqual "csv widened integer", "9999999999", _
        CStr(table.Rows.Item(1).Item("B"))
    AssertEqual "csv boolean column", False, table.Rows.Item(1).Item("C")
    AssertEqual "csv rolled date stays text", "2026-02-31", _
        CStr(table.Rows.Item(0).Item("D"))
    Set table = ROneCOne.Csv.DeserializeTable( _
        "V" & vbCrLf & """42""" & vbCrLf)
    AssertEqual "csv quoted number stays text", "42", _
        CStr(table.Rows.Item(0).Item("V"))

    mCurrentTest = "TestCsvSurface.Escapes"
    Set table = ROneCOne.DataTable("Notes")
    table.Column "Text", vbString
    table.LoadRow Array("line1" & vbCrLf & "line2")
    Set roundTripped = ROneCOne.Csv.DeserializeTable(table.ToCsv)
    AssertEqual "csv embedded newline", "line1" & vbCrLf & "line2", _
        CStr(roundTripped.Rows.Item(0).Item("Text"))

    mCurrentTest = "TestCsvSurface.Views"
    Set table = ROneCOne.DataTable("Scores")
    table.Column "Name", vbString
    table.Column "Score", vbLong
    table.LoadRow Array("Ada", 90)
    table.LoadRow Array("Grace", 95)
    Set view = ROneCOne.DataView(table) _
        .WithFilter(table.Rows!Score.AtLeast(95&))
    AssertEqual "csv view serializes filtered", _
        "Name,Score" & vbCrLf & "Grace,95" & vbCrLf, view.ToCsv

    mCurrentTest = "TestCsvSurface.FileComposition"
    filePath = ThisWorkbook.Path & "\ROneCOne_CsvTest.csv"
    ROneCOne.File.WriteAllText filePath, table.ToCsv
    Set roundTripped = ROneCOne.Csv.DeserializeTable( _
        ROneCOne.File.ReadAllText(filePath))
    AssertEqual "csv file composition", 2&, roundTripped.Rows.Count
    ROneCOne.File.Delete filePath

    mCurrentTest = "TestCsvSurface.Errors"
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable "A,B" & vbLf & "1"
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv ragged row rejected", ROneCOne.CsvError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable vbNullString
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv empty text rejected", ROneCOne.CsvError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable "A,a" & vbLf & "1,2"
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv duplicate header rejected", _
        ROneCOne.CsvError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable "A" & vbLf & "x""y"
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv stray quote rejected", ROneCOne.CsvError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable "A" & vbLf & """x""y"
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv text after closing quote rejected", _
        ROneCOne.CsvError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.DeserializeTable "A" & vbLf & "1", "T", "$.x"
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv array path rejected", _
        ROneCOne.InvalidArgumentError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.Serialize table, True
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv indented form rejected", _
        ROneCOne.InvalidArgumentError, csvErrorNumber
    csvErrorNumber = 0
    On Error Resume Next
    ROneCOne.Csv.Serialize 42
    csvErrorNumber = Err.Number
    On Error GoTo 0
    AssertEqual "csv scalar serialize rejected", _
        ROneCOne.CsvError, csvErrorNumber
End Sub

Private Sub TestProcessSurface()
    Dim capturedError As Long
    Dim result As ROneCOne
    Dim results As ROneCOne
    Dim source As ROneCOne
    Dim task As ROneCOne

    ' Awaitable shell commands: the process runs outside Excel while the
    ' Task polls its status; both output streams come back captured from
    ' scratch redirect files that the runtime deletes afterward.
    mCurrentTest = "TestProcessSurface.Echo"
    Set result = ROneCOne.Process.RunAsync("echo hello").Await
    AssertEqual "process echo exit code", 0&, result.ExitCode
    AssertTrue "process echo output", _
        InStr(1, result.StandardOutput, "hello") > 0
    AssertEqual "process echo stderr empty", vbNullString, _
        Trim$(result.StandardError)

    mCurrentTest = "TestProcessSurface.ExitCode"
    AssertEqual "process exit code passthrough", 7&, _
        ROneCOne.Process.RunAsync("exit 7").Await.ExitCode

    mCurrentTest = "TestProcessSurface.StandardError"
    Set result = ROneCOne.Process.RunAsync("echo oops 1>&2").Await
    AssertTrue "process stderr captured", _
        InStr(1, result.StandardError, "oops") > 0
    AssertEqual "process stdout empty", vbNullString, _
        Trim$(result.StandardOutput)

    mCurrentTest = "TestProcessSurface.MissingCommand"
    Set result = ROneCOne.Process.RunAsync( _
        "definitely_not_a_command_xyz").Await
    AssertTrue "process missing command fails", result.ExitCode <> 0
    AssertTrue "process missing command explains", _
        InStr(1, result.StandardError, "not recognized") > 0

    mCurrentTest = "TestProcessSurface.WorkingDirectory"
    Set result = ROneCOne.Process.RunAsync("cd", ThisWorkbook.Path).Await
    AssertTrue "process working directory", InStr(1, _
        result.StandardOutput, ThisWorkbook.Path, vbTextCompare) > 0

    mCurrentTest = "TestProcessSurface.WhenAll"
    Set results = ROneCOne.Task.WhenAll( _
        ROneCOne.Process.RunAsync("echo first"), _
        ROneCOne.Process.RunAsync("echo second")).Await
    AssertEqual "process whenall count", 2&, results.Count
    AssertTrue "process whenall first", _
        InStr(1, results.Item(0).StandardOutput, "first") > 0
    AssertTrue "process whenall second", _
        InStr(1, results.Item(1).StandardOutput, "second") > 0

    mCurrentTest = "TestProcessSurface.Cancellation"
    Set source = ROneCOne.CancellationTokenSource
    Set task = ROneCOne.Process.RunAsync( _
        "ping -n 6 127.0.0.1", , source.Token)
    source.Cancel
    capturedError = 0
    On Error Resume Next
    task.Await
    capturedError = Err.Number
    On Error GoTo 0
    AssertTrue "process canceled await raises", capturedError <> 0
    AssertTrue "process canceled task state", task.IsCanceled

    mCurrentTest = "TestProcessSurface.StandardInput"
    Set result = ROneCOne.Process.RunAsync("sort", , , _
        "banana" & vbCrLf & "apple" & vbCrLf & "cherry" & vbCrLf).Await
    AssertEqual "process stdin exit code", 0&, result.ExitCode
    AssertTrue "process stdin consumed", _
        InStr(1, result.StandardOutput, "apple") > 0 And _
        InStr(1, result.StandardOutput, "apple") < _
        InStr(1, result.StandardOutput, "banana")
    Set result = ROneCOne.Process.RunAsync("findstr b", , , _
        "banana" & vbCrLf & "apple" & vbCrLf & "berry" & vbCrLf).Await
    AssertTrue "process stdin filtered", _
        InStr(1, result.StandardOutput, "banana") > 0 And _
        InStr(1, result.StandardOutput, "apple") = 0
    ' With no input, stdin still closes, so a filter sees end-of-file
    ' immediately instead of waiting forever.
    AssertEqual "process closed stdin ends filter", 0&, _
        ROneCOne.Process.RunAsync("sort").Await.ExitCode

    mCurrentTest = "TestProcessSurface.Arguments"
    capturedError = 0
    On Error Resume Next
    ROneCOne.Process.RunAsync "   "
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "process empty command rejected", _
        ROneCOne.InvalidArgumentError, capturedError
    capturedError = 0
    On Error Resume Next
    Debug.Print ROneCOne.StandardOutput
    capturedError = Err.Number
    On Error GoTo 0
    AssertEqual "process result role guarded", _
        ROneCOne.InvalidOperationError, capturedError
End Sub

Private Sub TestRegexSurface()
    Dim expression As ROneCOne
    Dim firstMatch As ROneCOne
    Dim matches As ROneCOne
    Dim patternError As Long
    Dim pieces As ROneCOne

    ' The Regex surface wraps the in-box script engine expression object
    ' with the System.Text.RegularExpressions verbs; matches are typed
    ' values, so the whole LINQ surface applies to Matches results.
    mCurrentTest = "TestRegexSurface.Match"
    Set expression = ROneCOne.Regex("(\w+)@(\w+)\.com")
    AssertEqual "regex pattern", "(\w+)@(\w+)\.com", expression.Pattern
    AssertTrue "regex ismatch", expression.IsMatch("write ada@x.com today")
    AssertFalse "regex ismatch negative", expression.IsMatch("no emails")
    Set firstMatch = expression.Match("write ada@x.com today")
    AssertTrue "regex match success", firstMatch.Success
    AssertEqual "regex match value", "ada@x.com", firstMatch.Value
    AssertEqual "regex match index", 6&, firstMatch.FirstIndex
    AssertEqual "regex match length", 9&, firstMatch.Length
    AssertEqual "regex groups count", 3&, firstMatch.Groups.Count
    AssertEqual "regex group zero", "ada@x.com", _
        CStr(firstMatch.Groups.Item(0))
    AssertEqual "regex group one", "ada", CStr(firstMatch.Groups.Item(1))
    AssertEqual "regex group two", "x", CStr(firstMatch.Groups.Item(2))
    Set firstMatch = expression.Match("nothing here")
    AssertFalse "regex failed match", firstMatch.Success
    AssertEqual "regex failed value", vbNullString, firstMatch.Value

    mCurrentTest = "TestRegexSurface.MatchesReplaceSplit"
    Set matches = expression.Matches("ada@x.com and grace@y.com")
    AssertEqual "regex matches count", 2&, matches.Count
    AssertEqual "regex matches second", "grace@y.com", _
        matches.Item(1).Value
    AssertEqual "regex replace", "x:ada and y:grace", _
        expression.Replace("ada@x.com and grace@y.com", "$2:$1")
    Set pieces = ROneCOne.Regex("\s*,\s*").Split("a, b ,c,,d")
    AssertEqual "regex split count", 5&, pieces.Count
    AssertEqual "regex split piece", "c", CStr(pieces.Item(2))
    AssertEqual "regex split empty piece", vbNullString, _
        CStr(pieces.Item(3))
    Set pieces = ROneCOne.Regex("a*").Split("bab")
    AssertEqual "regex split skips empty matches", 2&, pieces.Count

    mCurrentTest = "TestRegexSurface.Flags"
    AssertEqual "regex flags", 2&, ROneCOne.Regex("^ada$", True, True) _
        .Matches("ADA" & vbLf & "ada").Count
    patternError = 0
    On Error Resume Next
    ROneCOne.Regex "(unclosed"
    patternError = Err.Number
    On Error GoTo 0
    AssertEqual "regex bad pattern raises", ROneCOne.RegexError, patternError
End Sub

Private Sub TestHashSurface()
    Dim conversionError As Long
    Dim decoded As Variant
    Dim digestHex As String
    Dim inputBytes As Variant

    ' Digests run through Windows CNG and are asserted against published
    ' vectors: FIPS 180 for the SHA family and MD5, RFC 4231 for HMAC.
    mCurrentTest = "TestHashSurface.KnownVectors"
    digestHex = ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256("abc"))
    AssertEqual "sha256 abc", _
        "BA7816BF8F01CFEA414140DE5DAE2223" & _
        "B00361A396177A9CB410FF61F20015AD", digestHex
    AssertEqual "sha256 empty", _
        "E3B0C44298FC1C149AFBF4C8996FB924" & _
        "27AE41E4649B934CA495991B7852B855", _
        ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256(vbNullString))
    AssertEqual "sha1 abc", "A9993E364706816ABA3E25717850C26C9CD0D89D", _
        ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha1("abc"))
    AssertEqual "md5 abc", "900150983CD24FB0D6963F7D28E17F72", _
        ROneCOne.Convert.ToHexString(ROneCOne.Hash.Md5("abc"))
    AssertEqual "sha512 abc length", 128&, _
        Len(ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha512("abc")))
    inputBytes = ROneCOne.Convert.FromHexString("616263")
    AssertEqual "sha256 bytes equal text", digestHex, _
        ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256(inputBytes))

    mCurrentTest = "TestHashSurface.Hmac"
    AssertEqual "hmac sha256 rfc4231 case 2", _
        "5BDCC146BF60754E6A042426089575C7" & _
        "5A003F089D2739839DEC58B964EC3843", _
        ROneCOne.Convert.ToHexString(ROneCOne.Hash.HmacSha256( _
        "Jefe", "what do ya want for nothing?"))

    mCurrentTest = "TestHashSurface.Convert"
    AssertEqual "base64 encode", "TWFu", ROneCOne.Convert.ToBase64String( _
        ROneCOne.Convert.FromHexString("4D616E"))
    decoded = ROneCOne.Convert.FromBase64String("TWFuTQ==")
    AssertEqual "base64 decode bound", 3&, CLng(UBound(decoded))
    AssertEqual "base64 decode value", 77&, CLng(decoded(3))
    AssertEqual "hex round trip", "4D616E", ROneCOne.Convert.ToHexString( _
        ROneCOne.Convert.FromHexString("4d616e"))
    conversionError = 0
    On Error Resume Next
    ROneCOne.Convert.FromBase64String "not base64!!"
    conversionError = Err.Number
    On Error GoTo 0
    AssertEqual "base64 rejects junk", _
        ROneCOne.InvalidArgumentError, conversionError
    conversionError = 0
    On Error Resume Next
    ROneCOne.Convert.FromHexString "ABC"
    conversionError = Err.Number
    On Error GoTo 0
    AssertEqual "hex rejects odd length", _
        ROneCOne.InvalidArgumentError, conversionError
End Sub

Private Sub TestDateTimeSurface()
    Dim currentUtc As ROneCOne
    Dim localWrap As ROneCOne
    Dim missError As Long
    Dim moment As ROneCOne
    Dim other As ROneCOne
    Dim parsedOk As Boolean
    Dim span As ROneCOne
    Dim tryResult As ROneCOne

    ' Instants mirror DateTimeOffset: a clock time plus an offset, stored as
    ' milliseconds since the Unix epoch in universal time. Local conversions
    ' are delegated to Windows so daylight saving applies per instant.
    mCurrentTest = "TestDateTimeSurface.Parse"
    Set moment = ROneCOne.DateTime.Parse("2026-07-24T18:30:05.123+02:00")
    AssertEqual "datetime year", 2026&, moment.Year
    AssertEqual "datetime month", 7&, moment.Month
    AssertEqual "datetime day", 24&, moment.Day
    AssertEqual "datetime hour", 18&, moment.Hour
    AssertEqual "datetime minute", 30&, moment.Minute
    AssertEqual "datetime second", 5&, moment.Second
    AssertEqual "datetime millisecond", 123&, moment.Millisecond
    AssertEqual "datetime offset minutes", 120#, moment.Offset.TotalMinutes
    AssertEqual "datetime iso round trip", _
        "2026-07-24T18:30:05.123+02:00", moment.ToIsoString
    AssertEqual "datetime unix seconds", 1784910605, _
        CDbl(moment.ToUnixTimeSeconds)
    AssertEqual "datetime unix milliseconds", 1784910605123#, _
        CDbl(moment.ToUnixTimeMilliseconds)
    AssertEqual "datetime universal view", "2026-07-24T16:30:05.123Z", _
        moment.ToUniversalTime.ToIsoString

    mCurrentTest = "TestDateTimeSurface.Epoch"
    Set moment = ROneCOne.DateTime.Parse("1970-01-01T00:00:00Z")
    AssertEqual "epoch unix zero", 0&, CDbl(moment.ToUnixTimeSeconds)
    AssertEqual "epoch day of week", 4&, moment.DayOfWeek
    AssertEqual "epoch day of year", 1&, moment.DayOfYear
    Set moment = ROneCOne.DateTime.FromUnixTimeMilliseconds(123)
    AssertEqual "from unix ms iso", "1970-01-01T00:00:00.123Z", _
        moment.ToIsoString
    Set moment = ROneCOne.DateTime.FromUnixTimeSeconds(1784910605)
    AssertEqual "from unix seconds round trip", 1784910605, _
        CDbl(moment.ToUnixTimeSeconds)

    mCurrentTest = "TestDateTimeSurface.Comparison"
    Set moment = ROneCOne.DateTime.Parse("2026-07-24T18:00:00+02:00")
    Set other = ROneCOne.DateTime.Parse("2026-07-24T16:00:00Z")
    AssertEqual "offset views same instant", 0&, moment.CompareTo(other)
    AssertEqual "later instant compares greater", 1&, _
        moment.AddSeconds(1).CompareTo(other)
    AssertEqual "earlier instant compares less", -1&, _
        moment.AddSeconds(-1).CompareTo(other)
    AssertEqual "subtract instants", 0#, moment.Subtract(other).TotalSeconds
    AssertEqual "to offset keeps instant", 0&, _
        moment.ToOffset(-300).CompareTo(other)
    AssertEqual "to offset clock hour", 11&, moment.ToOffset(-300).Hour

    mCurrentTest = "TestDateTimeSurface.Arithmetic"
    Set moment = ROneCOne.DateTime.Parse("2026-01-31T12:00:00Z")
    AssertEqual "add months clamps", "2026-02-28T12:00:00Z", _
        moment.AddMonths(1).ToIsoString
    AssertEqual "add years", "2027-01-31T12:00:00Z", _
        moment.AddYears(1).ToIsoString
    AssertEqual "add days fraction", "2026-02-01T00:00:00Z", _
        moment.AddDays(0.5).ToIsoString
    AssertEqual "add hours", "2026-01-31T13:30:00Z", _
        moment.AddHours(1.5).ToIsoString
    Set span = ROneCOne.TimeSpan.FromMinutes(90)
    AssertEqual "add timespan", "2026-01-31T13:30:00Z", _
        moment.Add(span).ToIsoString
    AssertEqual "subtract timespan", "2026-01-31T10:30:00Z", _
        moment.Subtract(span).ToIsoString

    mCurrentTest = "TestDateTimeSurface.Format"
    Set moment = ROneCOne.DateTime.Parse("2026-07-04T09:05:07.080Z")
    AssertEqual "custom tokens", "2026-07-04 09:05:07.080", _
        moment.ToString("yyyy-MM-dd HH:mm:ss.fff")
    AssertEqual "default tostring is iso", moment.ToIsoString, moment.ToString
    missError = 0
    On Error Resume Next
    moment.ToString "yyyy-QQ"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "unknown token raises", ROneCOne.FormatError, missError

    mCurrentTest = "TestDateTimeSurface.LocalZone"
    Set currentUtc = ROneCOne.DateTime.UtcNow
    AssertEqual "utcnow offset", 0#, currentUtc.Offset.TotalMinutes
    AssertTrue "now equals utcnow instant", _
        Abs(CDbl(ROneCOne.DateTime.Now.ToUnixTimeMilliseconds) - _
        CDbl(currentUtc.ToUnixTimeMilliseconds)) < 5000
    AssertTrue "vba now matches local clock", _
        Abs((ROneCOne.DateTime.Now.LocalDateTime - VBA.Now) * 86400#) < 5
    AssertEqual "today has no clock time", 0&, _
        ROneCOne.DateTime.Today.Hour * 3600 + _
        ROneCOne.DateTime.Today.Minute * 60 + ROneCOne.DateTime.Today.Second
    Set localWrap = ROneCOne.DateTime.FromLocal(DateSerial(2026, 1, 15))
    AssertEqual "fromlocal round trip", DateSerial(2026, 1, 15), _
        localWrap.LocalDateTime
    AssertEqual "fromlocal utc round trip", 0&, _
        ROneCOne.DateTime.FromUtc(localWrap.UtcDateTime).CompareTo(localWrap)
    Set moment = ROneCOne.DateTime.Parse("2026-01-15T00:00:00")
    AssertEqual "no offset assumes local", 0&, moment.CompareTo(localWrap)
    AssertEqual "tolocaltime keeps instant", 0&, _
        moment.ToLocalTime.CompareTo(moment)

    mCurrentTest = "TestDateTimeSurface.Failures"
    missError = 0
    On Error Resume Next
    ROneCOne.DateTime.Parse "2026-02-30"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "impossible date raises", ROneCOne.FormatError, missError
    missError = 0
    On Error Resume Next
    ROneCOne.DateTime.Parse "2026-07-24T25:00:00Z"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "impossible hour raises", ROneCOne.FormatError, missError
    missError = 0
    On Error Resume Next
    ROneCOne.DateTime.Parse "2026-07-24T10:00:00+15:00"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "impossible offset raises", ROneCOne.FormatError, missError
    parsedOk = ROneCOne.DateTime.TryParse("not a date", tryResult)
    AssertFalse "tryparse rejects junk", parsedOk
    AssertTrue "tryparse leaves nothing", tryResult Is Nothing
    parsedOk = ROneCOne.DateTime.TryParse("2026-07-24", tryResult)
    AssertTrue "tryparse accepts date", parsedOk
    AssertEqual "tryparse value", 2026&, tryResult.Year

    mCurrentTest = "TestDateTimeSurface.TimeSpan"
    Set span = ROneCOne.TimeSpan.FromMinutes(90)
    AssertEqual "timespan total hours", 1.5, span.TotalHours
    AssertEqual "timespan hours part", 1&, span.Hours
    AssertEqual "timespan minutes part", 30&, span.Minutes
    AssertEqual "timespan text", "01:30:00", span.ToString
    Set span = ROneCOne.TimeSpan.FromMinutes(-90)
    AssertEqual "negative hours part", -1&, span.Hours
    AssertEqual "negative minutes part", -30&, span.Minutes
    AssertEqual "negative text", "-01:30:00", span.ToString
    AssertEqual "duration text", "01:30:00", span.Duration.ToString
    AssertEqual "negate total minutes", 90#, span.Negate.TotalMinutes
    Set span = ROneCOne.TimeSpan.FromDays(1.5)
    AssertEqual "day text", "1.12:00:00", span.ToString
    AssertEqual "parse day round trip", 0&, _
        ROneCOne.TimeSpan.Parse("1.12:00:00").CompareTo(span)
    AssertEqual "parse fraction ms", 1500#, _
        ROneCOne.TimeSpan.Parse("00:00:01.5").TotalMilliseconds
    AssertEqual "add spans", 2.5, ROneCOne.TimeSpan.FromDays(1).Add( _
        ROneCOne.TimeSpan.FromDays(1.5)).TotalDays
    AssertEqual "subtract spans", -0.5, ROneCOne.TimeSpan.FromDays(1).Subtract( _
        ROneCOne.TimeSpan.FromDays(1.5)).TotalDays
    AssertEqual "zero total", 0#, ROneCOne.TimeSpan.Zero.TotalMilliseconds
    missError = 0
    On Error Resume Next
    ROneCOne.TimeSpan.Parse "25:00:00"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "timespan hour bound raises", ROneCOne.FormatError, missError

    mCurrentTest = "TestDateTimeSurface.Guards"
    missError = 0
    On Error Resume Next
    ROneCOne.DateTime.FromDays 1
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "factory role guard", _
        ROneCOne.InvalidOperationError, missError
    Set moment = ROneCOne.DateTime.Parse("2026-01-15T00:00:00Z")
    missError = 0
    On Error Resume Next
    moment.Add ROneCOne.DateTime.UtcNow
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "add requires timespan", _
        ROneCOne.TypeMismatchError, missError
End Sub

Private Sub TestStringsSurface()
    Dim builder As ROneCOne
    Dim firstGuid As String
    Dim maxSeen As Long
    Dim minSeen As Long
    Dim missError As Long
    Dim outOfRange As Long
    Dim randomBytes As Variant
    Dim rolled As Long
    Dim sample As Long
    Dim secondGuid As String

    ' Composite formatting mirrors String.Format with invariant text output:
    ' a period decimal separator and comma grouping on every machine.
    mCurrentTest = "TestStringsSurface.Format"
    AssertEqual "format basic", "Ada scored 95.5 points", _
        ROneCOne.Strings.Format("{0} scored {1} points", "Ada", 95.5)
    AssertEqual "format grouped number", "1,234,567.89", _
        ROneCOne.Strings.Format("{0:N2}", 1234567.891)
    AssertEqual "format fixed", "3.14", _
        ROneCOne.Strings.Format("{0:F2}", 3.14159)
    AssertEqual "format zero pad", "0042", _
        ROneCOne.Strings.Format("{0:D4}", 42)
    AssertEqual "format hex upper", "FF", _
        ROneCOne.Strings.Format("{0:X}", 255)
    AssertEqual "format hex padded", "00ff", _
        ROneCOne.Strings.Format("{0:x4}", 255)
    AssertEqual "format percent", "12.30%", _
        ROneCOne.Strings.Format("{0:P}", 0.123)
    AssertEqual "format alignment right", "   7", _
        ROneCOne.Strings.Format("{0,4}", 7)
    AssertEqual "format alignment left", "7   |", _
        ROneCOne.Strings.Format("{0,-4}|", 7)
    AssertEqual "format braces escape", "{7}", _
        ROneCOne.Strings.Format("{{{0}}}", 7)
    AssertEqual "format date tokens", "2026-07-24", ROneCOne.Strings.Format( _
        "{0:yyyy-MM-dd}", DateSerial(2026, 7, 24))
    AssertEqual "format date default", "2026-07-24T09:30:00", _
        ROneCOne.Strings.Format("{0}", _
        DateSerial(2026, 7, 24) + TimeSerial(9, 30, 0))
    AssertEqual "format datetime argument", "2026-07-24T16:30:05Z", _
        ROneCOne.Strings.Format("{0}", _
        ROneCOne.DateTime.Parse("2026-07-24T16:30:05Z"))
    AssertEqual "format timespan argument", "01:30:00", _
        ROneCOne.Strings.Format("{0}", ROneCOne.TimeSpan.FromMinutes(90))
    AssertEqual "format boolean", "True", _
        ROneCOne.Strings.Format("{0}", True)
    AssertEqual "format negative number", "-1,234.50", _
        ROneCOne.Strings.Format("{0:N2}", -1234.5)
    missError = 0
    On Error Resume Next
    ROneCOne.Strings.Format "{1}", "only one"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "format missing argument raises", _
        ROneCOne.FormatError, missError
    missError = 0
    On Error Resume Next
    ROneCOne.Strings.Format "{0:Q9}", 5
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "format unknown spec raises", ROneCOne.FormatError, missError

    mCurrentTest = "TestStringsSurface.StringBuilder"
    Set builder = ROneCOne.StringBuilder()
    builder.Append("Hello").Append(", ").Append "world"
    AssertEqual "builder text", "Hello, world", builder.ToString
    AssertEqual "builder length", 12&, builder.Length
    builder.AppendLine "!"
    builder.AppendFormat "{0:D3}", 7
    AssertEqual "builder composed", "Hello, world!" & vbCrLf & "007", _
        builder.ToString
    builder.Clear
    AssertEqual "builder cleared", 0&, builder.Length
    builder.Append 12.5
    AssertEqual "builder invariant number", "12.5", builder.ToString

    mCurrentTest = "TestStringsSurface.Guid"
    firstGuid = ROneCOne.Guid.NewGuid
    secondGuid = ROneCOne.Guid.NewGuid
    AssertTrue "guid shape", ROneCOne.Regex("^[0-9a-f]{8}-[0-9a-f]{4}-" & _
        "4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$").IsMatch(firstGuid)
    AssertTrue "guid unique", firstGuid <> secondGuid
    AssertEqual "guid empty", "00000000-0000-0000-0000-000000000000", _
        ROneCOne.Guid.EmptyGuid

    mCurrentTest = "TestStringsSurface.Random"
    randomBytes = ROneCOne.RandomNumberGenerator.GetBytes(16)
    AssertEqual "random bytes bound", 15&, CLng(UBound(randomBytes))
    AssertTrue "random bytes differ", _
        ROneCOne.Convert.ToHexString(randomBytes) <> _
        ROneCOne.Convert.ToHexString( _
        ROneCOne.RandomNumberGenerator.GetBytes(16))
    minSeen = 99
    maxSeen = -99
    outOfRange = 0
    For sample = 1 To 200
        rolled = ROneCOne.RandomNumberGenerator.GetInt32(1, 7)
        If rolled < minSeen Then minSeen = rolled
        If rolled > maxSeen Then maxSeen = rolled
        If rolled < 1 Or rolled > 6 Then outOfRange = outOfRange + 1
    Next sample
    AssertEqual "random int32 in range", 0&, outOfRange
    AssertTrue "random int32 spreads", maxSeen > minSeen
    missError = 0
    On Error Resume Next
    ROneCOne.RandomNumberGenerator.GetInt32 5, 5
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "random empty range raises", _
        ROneCOne.InvalidArgumentError, missError
End Sub

Private Sub TestEscapingSurface()
    Dim emoji As String
    Dim missError As Long

    ' Percent encoding runs over UTF-8 bytes with the RFC 3986 unreserved
    ' set; HTML encoding covers the five markup entities plus numeric
    ' references, decoding through surrogate pairs.
    mCurrentTest = "TestEscapingSurface.Uri"
    AssertEqual "uri escape basics", "a%20b%26c%3Dd", _
        ROneCOne.Uri.EscapeDataString("a b&c=d")
    AssertEqual "uri unreserved preserved", "AZaz09-._~", _
        ROneCOne.Uri.EscapeDataString("AZaz09-._~")
    AssertEqual "uri escape utf8", "%E2%98%83", _
        ROneCOne.Uri.EscapeDataString(ChrW$(9731))
    AssertEqual "uri unescape round trip", ChrW$(9731) & " x", _
        ROneCOne.Uri.UnescapeDataString("%E2%98%83%20x")
    AssertEqual "uri bad escape untouched", "%GG%2 a+b", _
        ROneCOne.Uri.UnescapeDataString("%GG%2 a+b")
    AssertEqual "uri empty text", vbNullString, _
        ROneCOne.Uri.EscapeDataString(vbNullString)

    mCurrentTest = "TestEscapingSurface.Html"
    AssertEqual "html encode markup", _
        "&lt;a &amp; &quot;b&quot;&#39;&gt;", _
        ROneCOne.WebUtility.HtmlEncode("<a & ""b""'>")
    AssertEqual "html encode non-ascii", "caf&#233;", _
        ROneCOne.WebUtility.HtmlEncode("caf" & ChrW$(233))
    AssertEqual "html decode entities", "<&AB&unknown;>", _
        ROneCOne.WebUtility.HtmlDecode("&lt;&amp;&#65;&#x42;&unknown;&gt;")
    emoji = ChrW$(&HD83D) & ChrW$(&HDE00)
    AssertEqual "html encode surrogate pair", "&#128512;", _
        ROneCOne.WebUtility.HtmlEncode(emoji)
    AssertEqual "html decode surrogate pair", emoji, _
        ROneCOne.WebUtility.HtmlDecode("&#128512;")
    AssertEqual "html decode leaves bare ampersand", "a & b", _
        ROneCOne.WebUtility.HtmlDecode("a & b")

    mCurrentTest = "TestEscapingSurface.Guards"
    missError = 0
    On Error Resume Next
    ROneCOne.Uri.HtmlEncode "x"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "escaping role guard", _
        ROneCOne.InvalidOperationError, missError
End Sub

Private Sub TestXmlSurface()
    Dim doc As ROneCOne
    Dim firstBook As ROneCOne
    Dim missError As Long
    Dim node As ROneCOne
    Dim table As ROneCOne
    Dim xmlPath As String
    Dim xmlText As String

    ' The Xml surface rides MSXML6 with its secure defaults intact: DTDs
    ' stay prohibited and external references are never resolved.
    mCurrentTest = "TestXmlSurface.Parse"
    xmlText = "<catalog created=""2026-01-05""><book id=""1"">" & _
        "<title>First</title><price>10.5</price></book>" & _
        "<book id=""2""><title>Second &amp; Third</title></book></catalog>"
    Set doc = ROneCOne.Xml.Parse(xmlText)
    AssertEqual "xml root name", "catalog", doc.Name
    AssertEqual "xml root attribute", "2026-01-05", _
        doc.GetAttribute("created")
    AssertTrue "xml has attribute", doc.HasAttribute("created")
    AssertFalse "xml missing attribute", doc.HasAttribute("missing")
    AssertEqual "xml elements count", 2&, doc.Elements.Count
    AssertEqual "xml elements filtered", 2&, doc.Elements("book").Count
    AssertEqual "xml select nodes", 2&, doc.SelectNodes("//book").Count
    Set node = doc.SelectSingleNode("//book[@id='2']/title")
    AssertEqual "xml predicate text", "Second & Third", node.Value
    AssertTrue "xml select miss is nothing", _
        doc.SelectSingleNode("//missing") Is Nothing
    Set firstBook = doc.Elements("book").Item(0)
    AssertEqual "xml child attribute", "1", firstBook.GetAttribute("id")
    AssertTrue "xml outer xml", _
        InStr(1, firstBook.OuterXml, "<title>First</title>") > 0
    AssertEqual "xml value concatenates", "First10.5", firstBook.Value

    mCurrentTest = "TestXmlSurface.Namespaces"
    Set doc = ROneCOne.Xml.Parse( _
        "<root xmlns=""urn:probe""><item/><item/></root>", _
        "xmlns:p='urn:probe'")
    AssertEqual "xml plain xpath misses namespace", 0&, _
        doc.SelectNodes("//item").Count
    AssertEqual "xml mapped xpath hits", 2&, _
        doc.SelectNodes("//p:item").Count

    mCurrentTest = "TestXmlSurface.Failures"
    missError = 0
    On Error Resume Next
    ROneCOne.Xml.Parse "<a><b></a>"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml bad markup raises", ROneCOne.XmlError, missError
    missError = 0
    On Error Resume Next
    ROneCOne.Xml.Parse "<!DOCTYPE a [<!ENTITY e ""x"">]><a>&e;</a>"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml dtd blocked", ROneCOne.XmlError, missError
    Set doc = ROneCOne.Xml.Parse("<a><b>1</b></a>")
    missError = 0
    On Error Resume Next
    doc.SelectNodes "///bad["
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml bad xpath raises", ROneCOne.XmlError, missError
    missError = 0
    On Error Resume Next
    doc.GetAttribute "missing"
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml absent attribute raises", ROneCOne.XmlError, missError

    mCurrentTest = "TestXmlSurface.Files"
    xmlPath = CStr(ROneCOne.Path.Combine( _
        CStr(ROneCOne.Path.GetTempPath), "ROneCOne_Xml_Test.xml"))
    ROneCOne.File.WriteAllText xmlPath, _
        "<?xml version=""1.0"" encoding=""utf-8""?><data><row>1</row></data>"
    Set doc = ROneCOne.Xml.Load(xmlPath)
    AssertEqual "xml file root", "data", doc.Name
    ROneCOne.File.Delete xmlPath
    missError = 0
    On Error Resume Next
    ROneCOne.Xml.Load xmlPath
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml missing file raises", ROneCOne.XmlError, missError

    mCurrentTest = "TestXmlSurface.Tables"
    Set table = ROneCOne.Xml.DeserializeTable(xmlText, "Books", "//book")
    AssertEqual "xml table name", "Books", table.TableName
    AssertEqual "xml table rows", 2&, table.Rows.Count
    AssertEqual "xml table columns", 3&, table.Columns.Count
    AssertEqual "xml typed id", 1&, table.Rows.Item(0).Item("id")
    AssertEqual "xml typed price", 10.5, table.Rows.Item(0).Item("price")
    AssertEqual "xml text column", "Second & Third", _
        table.Rows.Item(1).Item("title")
    AssertTrue "xml missing cell is null", _
        IsNull(table.Rows.Item(1).Item("price"))
    Set table = ROneCOne.DataTable("Pets")
    table.Column "name", vbString
    table.Column "age", vbLong
    table.LoadRow Array("Rex", 4)
    table.LoadRow Array("Mia & Co <3", Null)
    AssertEqual "xml serialize", "<NewDataSet><Pets><name>Rex</name>" & _
        "<age>4</age></Pets><Pets><name>Mia &amp; Co &lt;3</name>" & _
        "</Pets></NewDataSet>", table.ToXml
    Set table = ROneCOne.Xml.DeserializeTable(table.ToXml, "Round", "//Pets")
    AssertEqual "xml round trip rows", 2&, table.Rows.Count
    AssertEqual "xml round trip typed age", 4&, _
        table.Rows.Item(0).Item("age")

    mCurrentTest = "TestXmlSurface.Guards"
    missError = 0
    On Error Resume Next
    ROneCOne.Xml.Elements
    missError = Err.Number
    On Error GoTo 0
    AssertEqual "xml node role guard", _
        ROneCOne.InvalidOperationError, missError
End Sub

Private Function ElapsedSecondsSince(ByVal started As Double) As Double
    ElapsedSecondsSince = Timer - started
    If ElapsedSecondsSince < 0 Then
        ElapsedSecondsSince = ElapsedSecondsSince + 86400#
    End If
End Function

Private Sub ResetResults()
    mPassed = 0
    mFailed = 0
    mNextRow = 6
    With ThisWorkbook.Worksheets("Task and Data Results")
        .Range("A2:C400").ClearContents
        .Range("A2").Value2 = "Passed"
        .Range("A3").Value2 = "Failed"
        .Range("A4").Value2 = "Status"
        .Range("A5").Value2 = "Error"
    End With
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub

Private Sub AssertFalse(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, False, condition
End Sub

Private Sub AssertEqual( _
    ByVal testName As String, _
    ByVal expected As Variant, _
    ByVal actual As Variant _
)
    Dim passed As Boolean

    If IsObject(expected) Or IsObject(actual) Then
        passed = (TypeName(expected) = TypeName(actual))
    Else
        passed = (expected = actual)
    End If
    With ThisWorkbook.Worksheets("Task and Data Results")
        .Cells(mNextRow, 1).Value2 = testName
        .Cells(mNextRow, 2).Value2 = IIf(passed, "PASS", "FAIL")
        If Not passed Then
            .Cells(mNextRow, 3).Value2 = _
                "Expected " & CStr(expected) & ", got " & CStr(actual)
        End If
    End With
    mNextRow = mNextRow + 1
    If passed Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
    End If
End Sub
