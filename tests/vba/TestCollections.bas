Attribute VB_Name = "TestCollections"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mCurrentTest As String

Public Sub RunROneCOneCollectionTests()
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim capturedDescription As String

    On Error GoTo FatalFailure
    ResetResults

    mCurrentTest = "TestPrimitiveList"
    TestPrimitiveList
    mCurrentTest = "TestStrictElementType"
    TestStrictElementType
    mCurrentTest = "TestUserClassList"
    TestUserClassList
    mCurrentTest = "TestUserClassLinq"
    TestUserClassLinq
    mCurrentTest = "TestSyntaxSugar"
    TestSyntaxSugar
    mCurrentTest = "TestListMutation"
    TestListMutation
    mCurrentTest = "TestDeferredWhere"
    TestDeferredWhere
    mCurrentTest = "TestProjectionAndChaining"
    TestProjectionAndChaining
    mCurrentTest = "TestSequenceOperators"
    TestSequenceOperators
    mCurrentTest = "TestTerminals"
    TestTerminals
    mCurrentTest = "TestForEach"
    TestForEach
    mCurrentTest = vbNullString

    With ThisWorkbook.Worksheets("Collection Results")
        .Range("B2").Value2 = mPassed
        .Range("B3").Value2 = mFailed
        .Range("B4").Value2 = IIf(mFailed = 0, "PASS", "FAIL")
    End With
    Exit Sub

FatalFailure:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    With ThisWorkbook.Worksheets("Collection Results")
        .Range("B4").Value2 = "ERROR"
        .Range("B5").Value2 = mCurrentTest & " | " & CStr(capturedNumber) & _
            " | " & capturedSource & " | " & capturedDescription
    End With
End Sub

Private Sub TestSyntaxSugar()
    Dim actualError As Long
    Dim ada As GenericCustomer
    Dim customer As ROneCOne
    Dim customers As ROneCOne
    Dim experienced As ROneCOne
    Dim grace As GenericCustomer
    Dim invalid As ROneCOne
    Dim katherine As GenericCustomer
    Dim names As ROneCOne
    Dim numbers As ROneCOne
    Dim oldest As GenericCustomer
    Dim prototype As GenericCustomer
    Dim result As ROneCOne
    Dim value As ROneCOne

    Set prototype = New GenericCustomer
    Set customers = ROneCOne.ListOf(prototype)
    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40
    Set katherine = New GenericCustomer
    katherine.CustomerName = "Katherine"
    katherine.Age = 49
    customers.Add ada
    customers.Add grace
    customers.Add katherine

    Set customer = customers.Element
    Set experienced = customers.Where(customer("Age").AtLeast(CLng(40)))
    Set names = experienced _
        .Map(customer("CustomerName"), vbString) _
        .Sorted _
        .ToList
    Set oldest = customers.OrderByDescending(customer("Age")).First

    AssertEqual "sugar Where count", CLng(2), experienced.Count
    AssertEqual "sugar Map count", CLng(2), names.Count
    AssertEqual "sugar Sorted first", "Grace", names.Item(0)
    AssertEqual "sugar Sorted last", "Katherine", names.Item(1)
    AssertEqual "sugar object ordering", "Katherine", oldest.CustomerName
    AssertTrue "sugar Exists", customers.Exists( _
        customer("CustomerName").EqualTo("Ada"))
    AssertTrue "sugar All", customers.All(customer("Age").AtLeast(CLng(30)))
    AssertEqual "sugar aggregate", CDbl(44.5), CDbl( _
        experienced.Map(customer("Age"), vbLong).Average)

    Set numbers = ROneCOne.Range(CLng(1), CLng(4))
    Set value = numbers.Element
    Set result = numbers _
        .Where(value.AtLeast(CLng(2))) _
        .Map(value.Multiply(CLng(2)), vbLong) _
        .SortedDescending _
        .ToList
    AssertEqual "sugar primitive first", CLng(8), result.First
    AssertEqual "sugar primitive last", CLng(4), result.Last

    On Error Resume Next
    Set invalid = customers.Where(customer("Missing").EqualTo(1)).ToList
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "sugar missing member", ROneCOne.MemberAccessError, actualError
End Sub

Private Sub TestUserClassLinq()
    Dim ada As GenericCustomer
    Dim customers As ROneCOne
    Dim fixture As DelegateFixture
    Dim grace As GenericCustomer
    Dim names As ROneCOne
    Dim predicate As ROneCOne
    Dim prototype As GenericCustomer
    Dim selector As ROneCOne

    Set prototype = New GenericCustomer
    Set customers = ROneCOne.ListOf(prototype)
    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40
    customers.Add ada
    customers.Add grace

    Set fixture = New DelegateFixture
    Set predicate = ROneCOne.Func(fixture, "CustomerAtLeast40")
    Set predicate = predicate.Takes(prototype).Returns(vbBoolean)
    Set selector = ROneCOne.Func(fixture, "GetCustomerName")
    Set selector = selector.Takes(prototype).Returns(vbString)
    Set names = customers _
        .Where(predicate) _
        .SelectItems(selector, vbString) _
        .ToList

    AssertEqual "class LINQ count", CLng(1), names.Count
    AssertEqual "class LINQ projection", "Grace", names.Item(0)
End Sub

Public Sub RunROneCOneCollectionBenchmark()
    Dim elapsed As Double
    Dim filtered As ROneCOne
    Dim started As Double
    Dim values As ROneCOne
    Dim x As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    started = Timer
    Set values = ROneCOne.Range(CLng(1), CLng(10000))
    Set filtered = values _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .ToList
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#

    With ThisWorkbook.Worksheets("Collection Benchmarks")
        .Range("B2").Value2 = 10000
        .Range("B3").Value2 = elapsed
        .Range("B4").Value2 = filtered.Count
    End With
End Sub

Private Sub TestPrimitiveList()
    Dim numbers As ROneCOne

    mCurrentTest = "TestPrimitiveList.ListOf"
    Set numbers = ROneCOne.ListOf(vbLong)
    mCurrentTest = "TestPrimitiveList.Add1"
    numbers.Add CLng(5)
    mCurrentTest = "TestPrimitiveList.Add2"
    numbers.Add CLng(10)
    mCurrentTest = "TestPrimitiveList.Add3"
    numbers.Add CLng(15)

    mCurrentTest = "TestPrimitiveList.Count"
    AssertEqual "primitive count", CLng(3), numbers.Count
    mCurrentTest = "TestPrimitiveList.DefaultItem"
    AssertEqual "primitive index", CLng(5), numbers(CLng(0))
    mCurrentTest = "TestPrimitiveList.SetItem"
    numbers.Item(1) = CLng(11)
    mCurrentTest = "TestPrimitiveList.GetItem"
    AssertEqual "primitive setter", CLng(11), numbers.Item(1)
    mCurrentTest = "TestPrimitiveList.GenericTypeName"
    AssertEqual "generic name", "List<Long>", numbers.GenericTypeName
End Sub

Private Sub TestStrictElementType()
    Dim numbers As ROneCOne
    Dim actualError As Long

    Set numbers = ROneCOne.ListOf(vbLong)
    On Error Resume Next
    numbers.Add "not a Long"
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEqual "strict type error", ROneCOne.TypeMismatchError, actualError
    AssertEqual "failed add is atomic", CLng(0), numbers.Count
End Sub

Private Sub TestUserClassList()
    Dim prototype As GenericCustomer
    Dim ada As GenericCustomer
    Dim grace As GenericCustomer
    Dim returned As GenericCustomer
    Dim noCustomer As GenericCustomer
    Dim customers As ROneCOne
    Dim actualError As Long
    Dim wrongType As DelegateFixture

    Set prototype = New GenericCustomer
    Set customers = ROneCOne.ListOf(prototype)
    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40

    customers.Add ada
    customers.Add grace
    customers.Add noCustomer
    Set returned = customers(CLng(0))

    AssertTrue "class identity", returned Is ada
    AssertEqual "class property", "Grace", customers.Item(1).CustomerName
    AssertTrue "class Nothing", customers.Item(2) Is Nothing
    AssertEqual "class generic name", "List<GenericCustomer>", customers.GenericTypeName

    Set wrongType = New DelegateFixture
    On Error Resume Next
    customers.Add wrongType
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "class type error", ROneCOne.TypeMismatchError, actualError
End Sub

Private Sub TestListMutation()
    Dim numbers As ROneCOne
    Dim more As ROneCOne

    Set numbers = ROneCOne.ListOf(vbLong)
    Set more = ROneCOne.ListOf(vbLong)
    numbers.Add CLng(1)
    numbers.Add CLng(3)
    numbers.Insert CLng(1), CLng(2)
    more.Add CLng(4)
    more.Add CLng(5)
    numbers.AddRange more

    AssertEqual "insert", CLng(2), numbers.Item(1)
    AssertEqual "index of", CLng(3), numbers.IndexOf(CLng(4))
    AssertTrue "contains", numbers.Contains(CLng(5))
    AssertTrue "remove", numbers.Remove(CLng(3))
    numbers.RemoveAt CLng(0)
    AssertEqual "remove at", CLng(2), numbers.Item(0)
    numbers.Clear
    AssertEqual "clear", CLng(0), numbers.Count
End Sub

Private Sub TestDeferredWhere()
    Dim x As ROneCOne
    Dim numbers As ROneCOne
    Dim query As ROneCOne
    Dim filtered As ROneCOne

    Set numbers = ROneCOne.ListOf(vbLong)
    numbers.Add CLng(5)
    numbers.Add CLng(20)
    Set x = ROneCOne.Parameter(vbLong)
    Set query = numbers.Where(ROneCOne.Lambda(x.GreaterThan(CLng(10)), x))
    numbers.Add CLng(30)
    Set filtered = query.ToList

    AssertEqual "deferred count", CLng(2), filtered.Count
    AssertEqual "deferred first", CLng(20), filtered.Item(0)
    AssertEqual "deferred sees mutation", CLng(30), filtered.Item(1)
End Sub

Private Sub TestProjectionAndChaining()
    Dim x As ROneCOne
    Dim numbers As ROneCOne
    Dim projected As ROneCOne

    Set numbers = ROneCOne.Range(CLng(1), CLng(6)).ToList
    Set x = ROneCOne.Parameter(vbLong)
    Set projected = numbers _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .SelectItems(ROneCOne.Lambda(x.Multiply(CLng(10)), x), vbLong) _
        .OrderByDescending(ROneCOne.Lambda(x, x)) _
        .Take(CLng(2)) _
        .ToList

    AssertEqual "chain count", CLng(2), projected.Count
    AssertEqual "chain first", CLng(60), projected.Item(0)
    AssertEqual "chain second", CLng(40), projected.Item(1)
End Sub

Private Sub TestSequenceOperators()
    Dim values As ROneCOne
    Dim result As ROneCOne

    Set values = ROneCOne.ListOf(vbLong)
    values.Add CLng(2)
    values.Add CLng(2)
    values.Add CLng(3)
    Set result = values.Distinct _
        .Prepend(CLng(1)) _
        .Append(CLng(4)) _
        .Reverse _
        .Skip(CLng(1)) _
        .ToList

    AssertEqual "sequence count", CLng(3), result.Count
    AssertEqual "sequence first", CLng(3), result.Item(0)
    AssertEqual "sequence last", CLng(1), result.Item(2)
    AssertEqual "repeat", CLng(3), ROneCOne.Repeat("x", CLng(3)).Count
End Sub

Private Sub TestTerminals()
    Dim x As ROneCOne
    Dim values As ROneCOne

    Set values = ROneCOne.Range(CLng(1), CLng(5))
    Set x = ROneCOne.Parameter(vbLong)

    AssertTrue "any", values.AnyItem(ROneCOne.Lambda(x.EqualTo(CLng(3)), x))
    AssertTrue "all", values.All(ROneCOne.Lambda(x.GreaterThan(CLng(0)), x))
    AssertEqual "first", CLng(1), values.First
    AssertEqual "last", CLng(5), values.Last
    AssertEqual "sum", CDbl(15), CDbl(values.Sum)
    AssertEqual "average", CDbl(3), CDbl(values.Average)
    AssertEqual "min", CLng(1), values.Min
    AssertEqual "max", CLng(5), values.Max
End Sub

Private Sub TestForEach()
    Dim inner As Variant
    Dim nestedCount As Long
    Dim query As ROneCOne
    Dim queryCount As Long
    Dim queryNestedCount As Long
    Dim values As ROneCOne
    Dim value As Variant
    Dim total As Long
    Dim x As ROneCOne

    Set values = ROneCOne.Range(CLng(1), CLng(4))
    For Each value In values
        total = total + CLng(value)
    Next value

    For Each value In values
        For Each inner In values
            nestedCount = nestedCount + 1
        Next inner
    Next value

    Set x = ROneCOne.Parameter(vbLong)
    Set query = values.Where(ROneCOne.Lambda(x.GreaterThan(CLng(0)), x))
    For Each value In query
        For Each inner In query
            queryNestedCount = queryNestedCount + 1
        Next inner
    Next value
    values.Add CLng(5)
    For Each value In query
        queryCount = queryCount + 1
    Next value

    AssertEqual "For Each", CLng(10), total
    AssertEqual "nested For Each", CLng(16), nestedCount
    AssertEqual "nested query For Each", CLng(16), queryNestedCount
    AssertEqual "query enumeration refresh", CLng(5), queryCount
End Sub

Private Sub ResetResults()
    With ThisWorkbook.Worksheets("Collection Results")
        .Range("A2:C300").ClearContents
        .Range("A1:C1").Value = Array("Test", "Result", "Detail")
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
    With ThisWorkbook.Worksheets("Collection Results")
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
