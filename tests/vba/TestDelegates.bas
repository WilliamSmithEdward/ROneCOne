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
    mCurrentTest = "TestMethodDelegate"
    TestMethodDelegate
    mCurrentTest = "TestActionDelegate"
    TestActionDelegate
    mCurrentTest = "TestObjectReturn"
    TestObjectReturn
    mCurrentTest = "TestComposition"
    TestComposition
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

Private Sub TestMethodDelegate()
    Dim fixture As DelegateFixture
    Dim doubleValue As ROneCOne

    Set fixture = New DelegateFixture
    Set doubleValue = ROneCOne.FromMethod(fixture, "DoubleValue", 1)

    AssertEqual "method delegate", CLng(14), doubleValue(CLng(7))
End Sub

Private Sub TestActionDelegate()
    Dim fixture As DelegateFixture
    Dim recordValue As ROneCOne
    Dim ignored As Variant

    Set fixture = New DelegateFixture
    Set recordValue = ROneCOne.FromMethod(fixture, "RecordValue", 1)
    ignored = recordValue("captured")

    AssertEqual "action delegate", "captured", fixture.LastValue
End Sub

Private Sub TestObjectReturn()
    Dim fixture As DelegateFixture
    Dim echo As ROneCOne
    Dim returned As Object

    Set fixture = New DelegateFixture
    Set echo = ROneCOne.FromMethod(fixture, "EchoSelf", 0, True)
    Set returned = echo.Run()

    AssertTrue "object return", returned Is fixture
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
    Set doubleValue = ROneCOne.FromMethod(fixture, "DoubleValue", 1)
    Set pipeline = square.PipeTo(doubleValue)

    AssertEqual "composition", CLng(18), pipeline(CLng(3))
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
