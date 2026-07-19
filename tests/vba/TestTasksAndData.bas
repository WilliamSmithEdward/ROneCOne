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
    mCurrentTest = "TestAdvancedDataTable"
    TestAdvancedDataTable
    mCurrentTest = "TestDataViewAndMerge"
    TestDataViewAndMerge
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
    Dim taskValue As ROneCOne
    Dim work As ROneCOne

    Set work = ROneCOne.Value(CLng(42)).AsFunc
    Set taskValue = ROneCOne.Task.Run(work)

    AssertFalse "task starts incomplete", taskValue.IsCompleted
    AssertEqual "task await result", CLng(42), taskValue.Await
    AssertTrue "task completes", taskValue.IsCompleted
    AssertEqual "task status", "RanToCompletion", taskValue.Status
    AssertEqual "task result", CLng(42), taskValue.Result
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

    Set firstTask = ROneCOne.TaskFromResult(CLng(10))
    Set secondTask = ROneCOne.TaskFromResult(CLng(20))
    Set tasks = ROneCOne.ListOf(ROneCOne.Task, firstTask, secondTask)
    Set allTask = ROneCOne.Task.WhenAll(tasks)
    Set results = allTask.Await

    AssertEqual "WhenAll result count", CLng(2), results.Count
    AssertEqual "WhenAll preserves order", CLng(20), results.Item(1)
    AssertTrue "WhenAny completes", ROneCOne.Task.WhenAny(tasks).Await Is firstTask

    Set completion = ROneCOne.TaskCompletionSourceOf(vbLong)
    Set pending = completion.Task
    Set completed = ROneCOne.Task.FromResult(42&)
    Set tasks = ROneCOne.ListOf(ROneCOne.Task, pending, completed)
    Set winner = ROneCOne.Task.WhenAny(tasks).Await
    AssertTrue "WhenAny selects a completed task", winner Is completed
    AssertEqual "WhenAny completed result", 42&, winner.Result
End Sub

Private Sub TestCancellation()
    Dim source As ROneCOne
    Dim token As ROneCOne

    Set source = ROneCOne.CancellationTokenSource
    Set token = source.Token
    AssertFalse "token starts active", token.IsCancellationRequested
    source.Cancel
    AssertTrue "token observes cancellation", token.IsCancellationRequested
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
    Dim resultTask As ROneCOne
    Dim source As ROneCOne

    Set delayed = ROneCOne.Task.Delay(20&)
    AssertFalse "delay timeout", delayed.Wait(1&)
    AssertTrue "delay eventually completes", delayed.Wait(100&)

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
    Dim row As ROneCOne
    Dim table As ROneCOne

    Set table = ROneCOne.DataTable("Customers")
    table.AddColumn ROneCOne.DataColumn("Id", vbLong).AsUnique
    table.AddColumn ROneCOne.DataColumn("Name", vbString)
    Set table.PrimaryKey = table.Columns.Where("ColumnName").EqualTo("Id").ToList

    Set row = table.NewRow
    row.Item("Id") = CLng(1)
    row.Item("Name") = "Ada"
    table.AddRow row

    AssertEqual "table column count", CLng(2), table.Columns.Count
    AssertEqual "table row count", CLng(1), table.Rows.Count
    AssertEqual "row field by name", "Ada", table.Rows.Item(0).Item("Name")
    AssertEqual "primary-key find", "Ada", table.Find(CLng(1)).Item("Name")

    table.AcceptChanges
    AssertEqual "accepted row state", "Unchanged", row.RowState
    row.Item("Name") = "Augusta"
    AssertEqual "modified row state", "Modified", row.RowState
    row.RejectChanges
    AssertEqual "reject restores original", "Ada", row.Item("Name")

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
    Set relation = ROneCOne.DataRelation( _
        "CustomerOrders", customers.Columns.Item("Id"), _
        orders.Columns.Item("CustomerId"))
    data.AddRelation relation

    Set children = parent.GetChildRows("CustomerOrders")
    AssertEqual "relation child navigation", CLng(1), children.Count
    AssertTrue "relation parent navigation", _
        child.GetParentRow("CustomerOrders") Is parent
    AssertEqual "dataset table count", CLng(2), data.Tables.Count
End Sub

Private Sub TestAdvancedDataTable()
    Dim changes As ROneCOne
    Dim row As ROneCOne
    Dim selected As ROneCOne
    Dim table As ROneCOne

    Set table = ROneCOne.DataTable("People")
    table.Column("Id", vbLong).AutoNumber(100&, 10&).AsUnique
    table.Column("Name", vbString).WithDefault "Unknown"
    Set row = table.NewRow
    table.AddRow row

    AssertEqual "auto increment", 100&, row.Item("Id")
    AssertEqual "column default", "Unknown", row.Item("Name")
    Set selected = table.SelectRows(table.Rows!Name.EqualTo("Unknown"))
    AssertEqual "table SelectRows", 1&, selected.Count

    table.AcceptChanges
    row.Item("Name") = "Ada"
    Set changes = table.GetChanges
    AssertEqual "table GetChanges", 1&, changes.Rows.Count
    row.Delete
    AssertEqual "row Delete", "Deleted", row.RowState
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
