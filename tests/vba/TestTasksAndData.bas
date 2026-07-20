Attribute VB_Name = "TestTasksAndData"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mCurrentTest As String

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
    mCurrentTest = "TestRangeBridge"
    TestRangeBridge
    mCurrentTest = "TestRelationConstraints"
    TestRelationConstraints
    mCurrentTest = "TestExistingRelationValidation"
    TestExistingRelationValidation
    mCurrentTest = "TestProviderSurface"
    TestProviderSurface
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
    AssertFalse "provider reports native async truthfully", _
        connection.SupportsNativeAsync
    AssertEqual "provider async mode", "Cooperative", connection.AsyncMode
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
