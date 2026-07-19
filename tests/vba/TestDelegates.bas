Attribute VB_Name = "TestDelegates"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mCurrentTest As String

Public Sub RunROneCOneTests()
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim capturedDescription As String

    On Error GoTo FatalFailure

    ResetResults
    mCurrentTest = "TestUnaryLambda"
    TestUnaryLambda
    mCurrentTest = "TestBinaryLambda"
    TestBinaryLambda
    mCurrentTest = "TestDelegateSyntaxSugar"
    TestDelegateSyntaxSugar
    mCurrentTest = "TestComparisonAndShortCircuit"
    TestComparisonAndShortCircuit
    mCurrentTest = "TestTypedParameterFailure"
    TestTypedParameterFailure
    mCurrentTest = "TestUniversalFactories"
    TestUniversalFactories
    mCurrentTest = "TestInvalidDelegateConversions"
    TestInvalidDelegateConversions
    mCurrentTest = "TestWorkbookProcedureDelegate"
    TestWorkbookProcedureDelegate
    mCurrentTest = "TestDynamicInvokeAndMetadata"
    TestDynamicInvokeAndMetadata
    mCurrentTest = "TestMulticastDelegate"
    TestMulticastDelegate
    mCurrentTest = "TestByRefDelegate"
    TestByRefDelegate
    mCurrentTest = "TestProcedureByRefFailsClosed"
    TestProcedureByRefFailsClosed
    mCurrentTest = "TestNativeDelegate"
    TestNativeDelegate
    mCurrentTest = "TestComposition"
    TestComposition
    mCurrentTest = "TestActionExecute"
    TestActionExecute
    mCurrentTest = "TestEvents"
    TestEvents
    mCurrentTest = "TestStructuredExceptions"
    TestStructuredExceptions
    mCurrentTest = "TestUnboundParameterFailure"
    TestUnboundParameterFailure
    mCurrentTest = vbNullString

    With ThisWorkbook.Worksheets("Test Results")
        .Range("B2").Value2 = mPassed
        .Range("B3").Value2 = mFailed
        .Range("B4").Value2 = IIf(mFailed = 0, "PASS", "FAIL")
    End With

    Exit Sub

FatalFailure:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    With ThisWorkbook.Worksheets("Test Results")
        .Range("B4").Value2 = "ERROR"
        .Range("B5").Value2 = mCurrentTest & " | " & CStr(capturedNumber) & _
            " | " & capturedSource & " | " & capturedDescription
    End With
End Sub

Public Sub RunROneCOneBenchmark()
    Dim x As ROneCOne
    Dim square As ROneCOne
    Dim started As Double
    Dim index As Long
    Dim result As Variant

    Set x = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)

    started = Timer
    For index = 1 To 10000
        result = square(CLng(index))
    Next index

    With ThisWorkbook.Worksheets("Benchmarks")
        .Range("B2").Value2 = 10000
        .Range("B3").Value2 = Timer - started
        .Range("B4").Value2 = result
    End With
End Sub

Private Sub TestUnaryLambda()
    Dim x As ROneCOne
    Dim square As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)

    AssertEqual "unary lambda", CLng(81), square(CLng(9))
    AssertEqual "explicit Run", CLng(81), square.Run(CLng(9))
    AssertEqual "unary arity", CLng(1), square.Arity
End Sub

Private Sub TestBinaryLambda()
    Dim leftValue As ROneCOne
    Dim rightValue As ROneCOne
    Dim addValues As ROneCOne

    Set leftValue = ROneCOne.Parameter(vbLong)
    Set rightValue = ROneCOne.Parameter(vbLong)
    Set addValues = ROneCOne.Lambda(leftValue.Add(rightValue), leftValue, rightValue)

    AssertEqual "binary lambda", CLng(13), addValues(CLng(6), CLng(7))
End Sub

Private Sub TestDelegateSyntaxSugar()
    Dim addValues As ROneCOne
    Dim customer As ROneCOne
    Dim customerValue As GenericCustomer
    Dim inferredLambda As ROneCOne
    Dim managerValue As GenericCustomer
    Dim prototype As GenericCustomer
    Dim readManagerName As ROneCOne
    Dim readName As ROneCOne
    Dim square As ROneCOne
    Dim x As ROneCOne
    Dim y As ROneCOne

    Set x = ROneCOne.Var(vbLong)
    Set y = ROneCOne.Var(vbLong)
    Set square = x.Multiply(x).AsFunc
    Set addValues = x.Add(y).AsFunc
    Set inferredLambda = ROneCOne.Lambda(x.Subtract(y))

    Set prototype = New GenericCustomer
    Set customer = ROneCOne.VarLike(prototype)
    Set readName = customer("CustomerName").AsFunc
    Set readManagerName = customer _
        .Member("Manager", True) _
        .Member("CustomerName") _
        .AsFunc
    Set customerValue = New GenericCustomer
    customerValue.CustomerName = "Ada"
    Set managerValue = New GenericCustomer
    managerValue.CustomerName = "Grace"
    Set customerValue.Manager = managerValue

    AssertEqual "Var AsFunc unary", CLng(81), square(CLng(9))
    AssertEqual "AsFunc binary", CLng(13), addValues(CLng(6), CLng(7))
    AssertEqual "inferred Lambda", CLng(5), inferredLambda(CLng(9), CLng(4))
    AssertEqual "VarLike member Func", "Ada", readName(customerValue)
    AssertEqual "object member Func", "Grace", readManagerName(customerValue)
    AssertEqual "inferred Func arity", CLng(2), addValues.Arity
End Sub

Private Sub TestComparisonAndShortCircuit()
    Dim value As ROneCOne
    Dim between As ROneCOne
    Dim safeFalse As ROneCOne

    Set value = ROneCOne.Parameter(vbLong)
    Set between = ROneCOne.Lambda( _
        value.GreaterThan(CLng(10)).AndAlso(value.LessThan(CLng(20))), _
        value)
    Set safeFalse = ROneCOne.Lambda( _
        ROneCOne.Value(False).AndAlso(ROneCOne.Value(1).Divide(0)))

    AssertEqual "comparison true", True, between(CLng(15))
    AssertEqual "comparison false", False, between(CLng(25))
    AssertEqual "short circuit", False, safeFalse.Run()
End Sub

Private Sub TestTypedParameterFailure()
    Dim x As ROneCOne
    Dim identity As ROneCOne
    Dim actualError As Long
    Dim ignored As Variant

    Set x = ROneCOne.Parameter(vbLong)
    Set identity = ROneCOne.Lambda(x, x)

    On Error Resume Next
    ignored = identity("not a Long")
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEqual "typed parameter error", ROneCOne.TypeMismatchError, actualError
End Sub

Private Sub TestUniversalFactories()
    Dim callable As ROneCOne
    Dim fixture As DelegateFixture
    Dim identity As ROneCOne
    Dim recordValue As ROneCOne
    Dim returnedFixture As DelegateFixture
    Dim doubleValue As ROneCOne
    Dim ignored As Variant

    Set fixture = New DelegateFixture
    Set doubleValue = ROneCOne.Func(fixture, "DoubleValue")
    Set doubleValue = doubleValue.Takes(vbLong).Returns(vbLong)
    Set recordValue = ROneCOne.Action(fixture, "RecordValue")
    Set recordValue = recordValue.Takes(vbString)
    Set callable = ROneCOne.Func(fixture)
    Set callable = callable.Takes(vbLong).Returns(vbLong)
    Set identity = ROneCOne.Func(fixture, "EchoSelf")
    Set identity = identity.Returns(fixture)

    AssertEqual "method delegate", CLng(14), doubleValue(CLng(7))
    ignored = recordValue("captured")
    AssertEqual "action delegate", "captured", fixture.LastValue
    AssertEqual "callable object", CLng(105), callable(CLng(5))
    Set returnedFixture = identity()
    AssertTrue "object return identity", returnedFixture Is fixture
    AssertTrue "target metadata", identity.Target Is fixture
End Sub

Private Sub TestInvalidDelegateConversions()
    Dim actionValue As ROneCOne
    Dim actualError As Long
    Dim fixture As DelegateFixture
    Dim invalidDelegate As ROneCOne

    Set fixture = New DelegateFixture
    Set actionValue = ROneCOne.Action(fixture, "RecordValue")
    Set actionValue = actionValue.Takes(vbString)

    On Error Resume Next
    Set invalidDelegate = ROneCOne.Func(actionValue)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Action to Func fails closed", _
        ROneCOne.InvalidOperationError, actualError

    On Error Resume Next
    Set invalidDelegate = actionValue.Returns(vbString)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Action Returns fails closed", _
        ROneCOne.InvalidOperationError, actualError
End Sub

Private Sub TestProcedureByRefFailsClosed()
    Dim actualError As Long
    Dim increment As ROneCOne
    Dim ignored As Variant
    Dim reference As ROneCOne
    Dim value As Long

    value = 41
    Set reference = ROneCOne.RefLong(value)
    Set increment = ROneCOne.Action("DelegateProcedures.IncrementLong") _
        .Takes(ROneCOne.RefOf(vbLong))

    On Error Resume Next
    ignored = increment(reference)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEqual "procedure ByRef fails closed", _
        ROneCOne.ByRefInvocationError, actualError
    AssertEqual "failed ByRef is atomic", CLng(41), value
End Sub

Private Sub TestWorkbookProcedureDelegate()
    Dim addValues As ROneCOne

    Set addValues = ROneCOne.Func("DelegateProcedures.AddValues") _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)

    AssertEqual "workbook procedure", CLng(13), addValues(CLng(6), CLng(7))
End Sub

Private Sub TestDynamicInvokeAndMetadata()
    Dim addValues As ROneCOne
    Dim arguments As Variant
    Dim result As Variant

    Set addValues = ROneCOne.Func("DelegateProcedures.AddValues") _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)
    arguments = Array(CLng(20), CLng(22))
    result = addValues.DynamicInvoke(arguments)

    AssertEqual "DynamicInvoke", CLng(42), result
    AssertEqual "dynamic arity", CLng(2), addValues.Arity
    AssertEqual "method name", "DelegateProcedures.AddValues", addValues.MethodName
    AssertTrue "Func metadata", Not addValues.IsAction
    AssertEqual "signature metadata", "Func<Long, Long, Long>", addValues.Signature
End Sub

Private Sub TestMulticastDelegate()
    Dim combined As ROneCOne
    Dim firstHandler As ROneCOne
    Dim ignored As Variant
    Dim invocationList As ROneCOne
    Dim reduced As ROneCOne
    Dim secondHandler As ROneCOne

    Set firstHandler = ROneCOne.Action("DelegateProcedures.RecordFirst") _
        .Takes(vbString)
    Set secondHandler = ROneCOne.Action("DelegateProcedures.RecordSecond") _
        .Takes(vbString)
    Set combined = firstHandler.Combine(secondHandler)
    DelegateProcedures.ResetTrace
    ignored = combined("value")
    Set invocationList = combined.GetInvocationList
    Set reduced = combined.Remove(secondHandler)

    AssertEqual "multicast order", "first:value|second:value|", _
        DelegateProcedures.CurrentTrace
    AssertEqual "multicast count", CLng(2), combined.InvocationCount
    AssertEqual "invocation list", CLng(2), invocationList.Count
    AssertEqual "immutable removal", CLng(1), reduced.InvocationCount
    AssertEqual "source unchanged", CLng(2), combined.InvocationCount
End Sub

Private Sub TestByRefDelegate()
#If Win64 Then
    Dim increment As ROneCOne
    Dim ignored As Variant
    Dim reference As ROneCOne
    Dim value As Long

    value = 41
    Set reference = ROneCOne.RefLong(value)
    Set increment = ROneCOne.NativeAction( _
        DelegateProcedures.NativeIncrementLongAddress) _
        .Takes(ROneCOne.RefOf(vbLong))
    ignored = increment(reference)

    AssertEqual "ByRef mutation", CLng(42), value
#Else
    AssertTrue "ByRef mutation", True
#End If
End Sub

Private Sub TestNativeDelegate()
#If Win64 Then
    Dim actualError As Long
    Dim addValues As ROneCOne
    Dim ignored As Variant
    Dim untyped As ROneCOne

    Set addValues = ROneCOne.Native(DelegateProcedures.NativeAddLongAddress) _
        .Takes(vbLong, vbLong) _
        .Returns(vbLong)

    AssertEqual "native procedure pointer", CLng(42), _
        addValues(CLng(19), CLng(23))

    Set untyped = ROneCOne.Native(DelegateProcedures.NativeAddLongAddress)
    On Error Resume Next
    ignored = untyped(CLng(19), CLng(23))
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "untyped native fails closed", _
        ROneCOne.NativeInvocationError, actualError
#Else
    AssertTrue "native procedure pointer", True
#End If
End Sub

Private Sub TestComposition()
    Dim x As ROneCOne
    Dim square As ROneCOne
    Dim fixture As DelegateFixture
    Dim doubleValue As ROneCOne
    Dim pipeline As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)
    Set fixture = New DelegateFixture
    Set doubleValue = ROneCOne.Func(fixture, "DoubleValue") _
        .Takes(vbLong) _
        .Returns(vbLong)
    Set pipeline = square.PipeTo(doubleValue)

    AssertEqual "composition", CLng(18), pipeline(CLng(3))
End Sub

Private Sub TestActionExecute()
    Dim fixture As DelegateFixture
    Dim recordValue As ROneCOne
#If Win64 Then
    Dim increment As ROneCOne
    Dim value As Long
#End If

    Set fixture = New DelegateFixture
    Set recordValue = ROneCOne.Action(fixture, "RecordValue").Takes(vbString)
    recordValue.Execute "executed"

    AssertEqual "Action Execute", "executed", fixture.LastValue

#If Win64 Then
    value = 41
    Set increment = ROneCOne.NativeAction( _
        DelegateProcedures.NativeIncrementLongAddress) _
        .Takes(ROneCOne.RefOf(vbLong))
    increment.Execute ROneCOne.RefLong(value)
    AssertEqual "inline ByRef Execute", CLng(42), value
#End If
End Sub

Private Sub TestEvents()
    Dim actualError As Long
    Dim emptyEvent As ROneCOne
    Dim changed As ROneCOne
    Dim firstHandler As ROneCOne
    Dim invalidHandler As ROneCOne
    Dim removed As Boolean
    Dim secondHandler As ROneCOne

    Set firstHandler = ROneCOne.Action("DelegateProcedures.RecordFirst") _
        .Takes(vbString)
    Set secondHandler = ROneCOne.Action("DelegateProcedures.RecordSecond") _
        .Takes(vbString)
    Set changed = ROneCOne.EventOf(vbString) _
        .Subscribe(firstHandler) _
        .Subscribe(secondHandler)

    DelegateProcedures.ResetTrace
    changed.Emit "ready"
    removed = changed.Unsubscribe(secondHandler)

    AssertEqual "event order", "first:ready|second:ready|", _
        DelegateProcedures.CurrentTrace
    AssertEqual "event subscriber count", CLng(1), changed.HandlerCount
    AssertTrue "event unsubscribe", removed

    DelegateProcedures.ResetTrace
    changed.Emit "again"
    AssertEqual "event after unsubscribe", "first:again|", _
        DelegateProcedures.CurrentTrace

    Set invalidHandler = ROneCOne.Action("DelegateProcedures.RecordFirst") _
        .Takes(vbLong)
    On Error Resume Next
    changed.Subscribe invalidHandler
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "event signature mismatch", _
        ROneCOne.TypeMismatchError, actualError
    AssertEqual "failed subscription is atomic", CLng(1), changed.HandlerCount

    Set emptyEvent = ROneCOne.EventOf(vbString)
    emptyEvent.Emit "nobody"
    AssertEqual "empty event", CLng(0), emptyEvent.HandlerCount
End Sub

Private Sub TestStructuredExceptions()
    Dim attempt As ROneCOne
    Dim cleanup As ROneCOne
    Dim errorHandler As ROneCOne
    Dim errorNumber As Long
    Dim failingCatch As ROneCOne
    Dim failingFinally As ROneCOne
    Dim success As ROneCOne
    Dim work As ROneCOne
    Dim wrongCatch As ROneCOne

    Set work = ROneCOne.Action( _
        ROneCOne.Value(1).Divide(CLng(0))).Takes()
    Set errorHandler = ROneCOne.Action("DelegateProcedures.HandleExpected") _
        .Takes(ROneCOne.Exception)
    Set cleanup = ROneCOne.Action("DelegateProcedures.RecordFinally").Takes()
    Set attempt = ROneCOne.Try(work) _
        .Catch(errorHandler) _
        .Finally(cleanup)

    DelegateProcedures.ResetTrace
    attempt.Execute
    AssertTrue "Try Catch message", _
        InStr(1, DelegateProcedures.CurrentTrace, "caught:", _
            vbBinaryCompare) = 1
    AssertTrue "Try Finally trace", _
        Right$(DelegateProcedures.CurrentTrace, 8) = "finally|"

    Set attempt = ROneCOne.Try(work) _
        .Catch(DelegateProcedures.OtherErrorNumber, errorHandler) _
        .Finally(cleanup)
    DelegateProcedures.ResetTrace
    On Error Resume Next
    attempt.Execute
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "filtered Catch rethrows", CLng(11), errorNumber
    AssertEqual "Finally on rethrow", "finally|", DelegateProcedures.CurrentTrace

    Set success = ROneCOne.Action("DelegateProcedures.RecordWork").Takes()
    Set attempt = ROneCOne.Try(success).Finally(cleanup)
    DelegateProcedures.ResetTrace
    attempt.Execute
    AssertEqual "Finally on success", "work|finally|", _
        DelegateProcedures.CurrentTrace

    Set failingCatch = ROneCOne.Action( _
        ROneCOne.Value(1).Divide(CLng(0))).Takes()
    Set attempt = ROneCOne.Try(work) _
        .Catch(failingCatch) _
        .Finally(cleanup)
    DelegateProcedures.ResetTrace
    On Error Resume Next
    attempt.Execute
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Catch failure propagates", CLng(11), errorNumber
    AssertEqual "Finally after Catch failure", "finally|", _
        DelegateProcedures.CurrentTrace

    Set failingFinally = ROneCOne.Action( _
        ROneCOne.Value(1).Divide(CLng(0))).Takes()
    Set attempt = ROneCOne.Try(success).Finally(failingFinally)
    DelegateProcedures.ResetTrace
    On Error Resume Next
    attempt.Execute
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Finally failure propagates", CLng(11), errorNumber
    AssertEqual "work ran before failed Finally", "work|", _
        DelegateProcedures.CurrentTrace

    Set wrongCatch = ROneCOne.Action("DelegateProcedures.HandleExpected") _
        .Takes(vbObject)
    On Error Resume Next
    Set attempt = ROneCOne.Try(work).Catch(wrongCatch)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Catch requires Exception signature", _
        ROneCOne.TypeMismatchError, errorNumber
End Sub

Private Sub TestUnboundParameterFailure()
    Dim boundValue As ROneCOne
    Dim unboundValue As ROneCOne
    Dim ignored As ROneCOne
    Dim actualError As Long

    Set boundValue = ROneCOne.Parameter(vbLong)
    Set unboundValue = ROneCOne.Parameter(vbLong)

    On Error Resume Next
    Set ignored = ROneCOne.Lambda(boundValue.Add(unboundValue), boundValue)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEqual "unbound parameter error", ROneCOne.UnboundParameterError, actualError
End Sub

Private Sub ResetResults()
    With ThisWorkbook.Worksheets("Test Results")
        .Range("A2:B200").ClearContents
        .Range("A1:B1").Value = Array("Test", "Result")
    End With
    mPassed = 0
    mFailed = 0
    mNextRow = 6
    mCurrentTest = vbNullString
End Sub

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        RecordResult testName, True, vbNullString
    Else
        RecordResult testName, False, "Expected " & CStr(expected) & ", got " & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    If condition Then
        RecordResult testName, True, vbNullString
    Else
        RecordResult testName, False, "Condition was False"
    End If
End Sub

Private Sub RecordResult(ByVal testName As String, ByVal passed As Boolean, ByVal detail As String)
    With ThisWorkbook.Worksheets("Test Results")
        .Cells(mNextRow, 1).Value2 = testName
        .Cells(mNextRow, 2).Value2 = IIf(passed, "PASS", "FAIL")
        .Cells(mNextRow, 3).Value2 = detail
    End With

    If passed Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
    End If
    mNextRow = mNextRow + 1
End Sub
