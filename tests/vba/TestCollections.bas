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
    mCurrentTest = "TestNumericWidening"
    TestNumericWidening
    mCurrentTest = "TestIndexedAccessScales"
    TestIndexedAccessScales
    mCurrentTest = "TestKeyedMutationConsistency"
    TestKeyedMutationConsistency
    mCurrentTest = "TestListMirrorConsistency"
    TestListMirrorConsistency
    mCurrentTest = "TestUserClassList"
    TestUserClassList
    mCurrentTest = "TestUserClassLinq"
    TestUserClassLinq
    mCurrentTest = "TestSyntaxSugar"
    TestSyntaxSugar
    mCurrentTest = "TestPredicateSystem"
    TestPredicateSystem
    mCurrentTest = "TestCollectionInitializers"
    TestCollectionInitializers
    mCurrentTest = "TestCollectionActions"
    TestCollectionActions
    mCurrentTest = "TestListMutation"
    TestListMutation
    mCurrentTest = "TestDeferredWhere"
    TestDeferredWhere
    mCurrentTest = "TestProjectionAndChaining"
    TestProjectionAndChaining
    mCurrentTest = "TestOrderingSystem"
    TestOrderingSystem
    mCurrentTest = "TestSequenceOperators"
    TestSequenceOperators
    mCurrentTest = "TestCompleteLinqSurface"
    TestCompleteLinqSurface
    mCurrentTest = "TestModernLinqSurface"
    TestModernLinqSurface
    mCurrentTest = "TestTerminals"
    TestTerminals
    mCurrentTest = "TestForEach"
    TestForEach
    mCurrentTest = "TestExpressionDisplay"
    TestExpressionDisplay
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

Private Sub TestPredicateSystem()
    Dim actualError As Long
    Dim ada As GenericCustomer
    Dim allowedCities As ROneCOne
    Dim City As Variant
    Dim comparer As ROneCOne
    Dim customers As ROneCOne
    Dim equalityComparer As ROneCOne
    Dim grace As GenericCustomer
    Dim katherine As GenericCustomer
    Dim noCustomer As GenericCustomer
    Dim reportPredicate As ROneCOne
    Dim result As ROneCOne
    Dim strings As ROneCOne

    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    ada.City = "London"
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40
    grace.City = "Arlington"
    Set grace.Manager = ada
    Set katherine = New GenericCustomer
    katherine.CustomerName = "Katherine"
    katherine.Age = 49
    katherine.City = "Cleveland"
    Set katherine.Manager = grace
    Set customers = ROneCOne.ListFrom(ada, grace, katherine)

    Set allowedCities = ROneCOne.ListOf( _
        vbString, "London", "Cleveland")
    Set result = customers.Where("City").IsIn(allowedCities).ToList
    AssertEqual "membership contextual", CLng(2), result.Count
    Set result = customers.Where("City") _
        .OneOf(Array("London", "Cleveland")).ToList
    AssertEqual "OneOf array", CLng(2), result.Count
    Set result = customers.Where(allowedCities.Contains(customers!City)).ToList
    AssertEqual "collection.Contains expression", CLng(2), result.Count
    Set result = customers.Where( _
        allowedCities.ContainsMember(customers, "City")).ToList
    AssertEqual "ContainsMember expression", CLng(2), result.Count

    AssertEqual "ignore-case equality", CLng(1), _
        customers.Where("City").EqualToIgnoreCase("LONDON").Count
    AssertEqual "ignore-case contains", CLng(1), _
        customers.Where("CustomerName").ContainsIgnoreCase("THER").Count
    AssertEqual "ignore-case pattern", CLng(1), _
        customers.Where("CustomerName") _
            .MatchesPatternIgnoreCase("k*").Count
    AssertEqual "null-safe nested path", CLng(1), _
        customers.Where("Manager?.Age").AtLeast(CLng(40)).Count
    AssertEqual "null-safe path short-circuits remainder", CLng(1), _
        customers.Where("Manager?.Manager.Age").EqualTo(CLng(36)).Count

    AssertEqual "Both composition", CLng(1), customers.Where( _
        customers.Condition("Age").AtLeast(CLng(40)).Both( _
            customers.Condition("City").EqualTo("Arlington"))).Count
    AssertEqual "Either composition", CLng(2), customers.Where( _
        customers.Match("City", "London").Either( _
            customers.Match("City", "Cleveland"))).Count
    AssertEqual "Negated composition", CLng(2), customers.Where( _
        customers.Match("City", "London").Negated).Count
    AssertEqual "WhereNot", CLng(2), _
        customers.WhereNot(customers.Match("City", "London")).Count
    AssertEqual "Always", CLng(3), customers.Count(customers.Always)
    AssertTrue "Never", customers.None(customers.Never)

    AssertEqual "predicate Count", CLng(2), _
        customers.Count(customers.Condition("Age").AtLeast(CLng(40)))
    AssertEqual "FirstOrDefault value", CLng(0), _
        ROneCOne.ListOf(vbLong).FirstOrDefault
    AssertEqual "LastOrDefault value", CLng(0), _
        ROneCOne.ListOf(vbLong).LastOrDefault
    Set noCustomer = customers.FirstOrDefault(customers.Match("Age", CLng(99)))
    AssertTrue "FirstOrDefault object", noCustomer Is Nothing
    Set grace = customers.SingleItem(customers.Match("Age", CLng(40)))
    AssertEqual "Single", "Grace", grace.CustomerName
    Set noCustomer = customers.SingleOrDefault(customers.Match("Age", CLng(99)))
    AssertTrue "SingleOrDefault object", noCustomer Is Nothing
    On Error Resume Next
    Set noCustomer = customers.SingleItem
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Single rejects many", _
        ROneCOne.InvalidOperationError, actualError

    Set equalityComparer = ROneCOne.EqualityComparer( _
        "DelegateProcedures.TextEqualsIgnoreCase")
    Set comparer = ROneCOne.Comparer( _
        "DelegateProcedures.CompareTextIgnoreCase")
    Set strings = ROneCOne.ListOf( _
        vbString, "Ada", "ADA", "grace")
    AssertEqual "custom Distinct comparer", CLng(2), _
        strings.Distinct(equalityComparer).Count
    AssertTrue "custom Contains comparer", _
        strings.Contains("ada", equalityComparer)
    Set result = strings.Where( _
        ROneCOne.ListOf(vbString, "ada").Contains( _
            strings.Element, equalityComparer)).ToList
    AssertEqual "custom predicate membership comparer", CLng(2), result.Count
    AssertEqual "custom ordering comparer", "Ada", _
        strings.Order(comparer).First
    AssertTrue "SequenceEqual comparer", _
        ROneCOne.ListOf(vbString, "ada", "Grace").SequenceEqual( _
            Array("ADA", "grace"), equalityComparer)

    Set ada.Reports = ROneCOne.ListFrom(grace, katherine)
    Set grace.Reports = ROneCOne.ListFrom(ada)
    Set katherine.Reports = ROneCOne.ListLike(ada)
    Set reportPredicate = ada.Reports.Condition("Age").AtLeast(CLng(40))
    AssertEqual "nested Any", CLng(1), _
        customers.WhereAny("Reports", reportPredicate).Count
    AssertEqual "nested All", CLng(3), customers.WhereAll( _
        "Reports", ada.Reports.Condition("Age").AtLeast(CLng(36))).Count
    AssertEqual "nested None", CLng(2), _
        customers.WhereNone("Reports", reportPredicate).Count
    AssertEqual "fluent nested Any", CLng(1), _
        customers.Where("Reports").AnyMatch(reportPredicate).Count
End Sub

Private Sub TestCollectionInitializers()
    Dim actualError As Long
    Dim ada As GenericCustomer
    Dim customers As ROneCOne
    Dim grace As GenericCustomer
    Dim inferredEmpty As ROneCOne
    Dim numbers As ROneCOne
    Dim values As Collection

    Set numbers = ROneCOne.ListOf( _
        vbLong, CLng(1), CLng(2), CLng(3))
    AssertEqual "typed initializer count", CLng(3), numbers.Count
    AssertEqual "typed initializer value", CLng(2), numbers.Item(1)

    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    Set customers = ROneCOne.ListFrom(ada, grace)
    Set inferredEmpty = ROneCOne.ListLike(ada)

    AssertEqual "inferred initializer type", _
        "List<GenericCustomer>", customers.GenericTypeName
    AssertTrue "inferred initializer identity", customers.Item(0) Is ada
    AssertEqual "ListLike type", _
        "List<GenericCustomer>", inferredEmpty.GenericTypeName
    AssertEqual "ListLike empty", CLng(0), inferredEmpty.Count

    Set values = New Collection
    values.Add CLng(4)
    values.Add CLng(5)
    numbers.AddRange values
    numbers.AddRange Array(CLng(6), CLng(7))
    AssertEqual "AddRange Collection and array", CLng(7), numbers.Count

    On Error Resume Next
    numbers.AddRange Array(CLng(8), "not a Long")
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "AddRange strict type", ROneCOne.TypeMismatchError, actualError
    AssertEqual "AddRange remains atomic", CLng(7), numbers.Count
End Sub

Private Sub TestCollectionActions()
    Dim actionValue As ROneCOne
    Dim emptyValues As ROneCOne
    Dim evenValues As ROneCOne
    Dim predicate As ROneCOne
    Dim values As ROneCOne

    Set values = ROneCOne.ListOf( _
        vbLong, CLng(1), CLng(2), CLng(3), CLng(4))
    Set actionValue = ROneCOne.Action("DelegateProcedures.AccumulateLong") _
        .Takes(vbLong)
    DelegateProcedures.ResetTotal
    values.ForEach actionValue

    AssertEqual "ForEach Action", CLng(10), DelegateProcedures.CurrentTotal
    AssertTrue "Exists without predicate", values.Exists
    Set emptyValues = ROneCOne.ListOf(vbLong)
    AssertTrue "Exists empty", Not emptyValues.Exists
    AssertEqual "JoinText", "1|2|3|4", values.JoinText("|")

    Set predicate = ROneCOne.Func("DelegateProcedures.IsEvenLong") _
        .Takes(vbLong) _
        .Returns(vbBoolean)
    Set evenValues = values.Where(predicate).ToList
    AssertEqual "procedure Func in LINQ", CLng(2), evenValues.Count
End Sub

Private Sub TestSyntaxSugar()
    Dim Age As Variant
    Dim actualError As Long
    Dim ada As GenericCustomer
    Dim customers As ROneCOne
    Dim distinctCities As ROneCOne
    Dim experienced As ROneCOne
    Dim grace As GenericCustomer
    Dim invalid As ROneCOne
    Dim katherine As GenericCustomer
    Dim names As ROneCOne
    Dim numbers As ROneCOne
    Dim oldest As GenericCustomer
    Dim predicate As ROneCOne
    Dim prototype As GenericCustomer
    Dim result As ROneCOne
    Dim value As ROneCOne
    Dim youngest As GenericCustomer

    Set prototype = New GenericCustomer
    Set customers = ROneCOne.ListOf(prototype)
    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    ada.Active = True
    ada.City = "London"
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40
    grace.Active = True
    grace.City = "London"
    Set grace.Manager = ada
    Set katherine = New GenericCustomer
    katherine.CustomerName = "Katherine"
    katherine.Age = 49
    katherine.Active = False
    katherine.City = vbNullString
    Set katherine.Manager = grace
    customers.Add ada
    customers.Add grace
    customers.Add katherine

    Set experienced = customers.Where("Age").AtLeast(CLng(40))
    Set names = experienced _
        .Map("CustomerName", vbString) _
        .Order _
        .ToList
    Set oldest = customers.OrderByDescending("Age").First
    Set youngest = customers.MinBy("Age")
    Set distinctCities = customers.DistinctBy("City").ToList

    AssertEqual "sugar Where count", CLng(2), experienced.Count
    AssertEqual "sugar Map count", CLng(2), names.Count
    AssertEqual "sugar Order first", "Grace", names.Item(0)
    AssertEqual "sugar Order last", "Katherine", names.Item(1)
    AssertEqual "sugar object ordering", "Katherine", oldest.CustomerName
    AssertEqual "sugar MinBy", "Ada", youngest.CustomerName
    AssertEqual "sugar MaxBy", "Katherine", _
        customers.MaxBy("Age").CustomerName
    AssertEqual "sugar DistinctBy", CLng(2), distinctCities.Count
    AssertTrue "sugar Exists", customers.Exists( _
        customers.Condition("CustomerName").EqualTo("Ada"))
    AssertTrue "sugar All", customers.All( _
        customers.Condition("Age").AtLeast(CLng(30)))
    AssertEqual "sugar aggregate", CDbl(44.5), _
        CDbl(experienced.Average("Age"))
    AssertEqual "sugar Sum selector", CLng(125), CLng(customers.Sum("Age"))
    AssertEqual "sugar JoinText selector", "Ada|Grace|Katherine", _
        customers.JoinText("|", "CustomerName")

    Set result = customers.Where("Age").Between(CLng(36), CLng(40)).ToList
    AssertEqual "sugar Between", CLng(2), result.Count
    Set result = customers.Where("CustomerName") _
        .OneOf("Ada", "Katherine").ToList
    AssertEqual "sugar OneOf", CLng(2), result.Count
    Set result = customers.Where("CustomerName").StartsWith("G").ToList
    AssertEqual "sugar StartsWith", CLng(1), result.Count
    Set result = customers.Where("CustomerName").EndsWith("e").ToList
    AssertEqual "sugar EndsWith", CLng(2), result.Count
    Set result = customers.Where("CustomerName") _
        .ContainsText("ther").ToList
    AssertEqual "sugar ContainsText", CLng(1), result.Count
    Set result = customers.Where("CustomerName").Contains("ther").ToList
    AssertEqual "C#-style Contains", CLng(1), result.Count
    Set result = customers.Where("CustomerName") _
        .MatchesPattern("K*").ToList
    AssertEqual "sugar MatchesPattern", CLng(1), result.Count
    Set result = customers.Where("City").IsNullOrEmpty.ToList
    AssertEqual "sugar IsNullOrEmpty", CLng(1), result.Count
    Set result = customers.Where("Active").IsTrue.ToList
    AssertEqual "sugar IsTrue", CLng(2), result.Count
    Set result = customers.Where("Active").IsFalse.ToList
    AssertEqual "sugar IsFalse", CLng(1), result.Count
    Set result = customers.Where("Manager").IsNothing.ToList
    AssertEqual "sugar IsNothing", CLng(1), result.Count
    Set result = customers.Where("Manager").IsNotNothing.ToList
    AssertEqual "sugar IsNotNothing", CLng(2), result.Count
    Set result = customers _
        .Where("Manager").IsNotNothing _
        .Where("Manager.Age").AtLeast(CLng(40)) _
        .ToList
    AssertEqual "sugar nested path", CLng(1), result.Count

    Set result = customers.Where( _
        customers.Condition("Age").Between(CLng(35), CLng(45)) _
        .AndAlso(customers.Condition("City").EqualTo("London"))) _
        .ToList
    AssertEqual "sugar composed conditions", CLng(2), result.Count

    Set predicate = customers.Predicate( _
        "DelegateProcedures.IsExperiencedCustomer")
    AssertEqual "inferred predicate signature", _
        "Func<GenericCustomer, Boolean>", predicate.Signature
    Set result = customers.WhereMethod( _
        "DelegateProcedures.IsExperiencedCustomer").ToList
    AssertEqual "inferred procedure predicate", CLng(2), result.Count

    On Error Resume Next
    Set predicate = customers.Predicate( _
        ROneCOne.Func("DelegateProcedures.IsExperiencedCustomer") _
            .Takes(vbLong) _
            .Returns(vbBoolean))
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "predicate type mismatch fails early", _
        ROneCOne.TypeMismatchError, actualError

    Set result = customers.Where(customers!Age.AtLeast(CLng(40))).ToList
    AssertEqual "native bang member", CLng(2), result.Count
    With customers
        Set result = .Where(!Age.AtLeast(CLng(40))).ToList
    End With
    AssertEqual "With bang member", CLng(2), result.Count

    Set numbers = ROneCOne.Range(CLng(1), CLng(4))
    Set value = numbers.Element
    Set result = numbers _
        .Where(value.AtLeast(CLng(2))) _
        .Map(value.Multiply(CLng(2)), vbLong) _
        .OrderDescending _
        .ToList
    AssertEqual "sugar primitive first", CLng(8), result.First
    AssertEqual "sugar primitive last", CLng(4), result.Last

    On Error Resume Next
    Set invalid = customers.Where("Missing").EqualTo(1).ToList
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
    Set names = customers.WhereMethod(fixture, "CustomerAtLeast40") _
        .Map("CustomerName", vbString) _
        .ToList
    AssertEqual "inferred object-method predicate", "Grace", names.Item(0)
End Sub

Public Sub RunROneCOneCollectionBenchmark()
    Dim capacity As Long
    Dim dictionary As ROneCOne
    Dim elapsed As Double
    Dim filtered As ROneCOne
    Dim ordered As ROneCOne
    Dim orderingElapsed As Double
    Dim started As Double
    Dim values As ROneCOne
    Dim x As ROneCOne
    Dim genericReadCheck As Long
    Dim genericReadElapsed As Double
    Dim hashElapsed100K As Double
    Dim hashElapsed10K As Double
    Dim hashSet As ROneCOne
    Dim index As Long
    Dim lastValue As Long
    Dim listWriteCheck As Long
    Dim listWriteElapsed As Double
    Dim mutationElapsed As Double
    Dim rowsLoopElapsed As Double
    Dim rowsTotal As Long
    Dim table As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    started = Timer
    Set values = ROneCOne.Range(CLng(1), CLng(10000))
    Set filtered = values _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .ToList
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#

    started = Timer
    Set ordered = values _
        .OrderBy(ROneCOne.Lambda(x.Modulo(CLng(100)), x)) _
        .ThenByDescending(ROneCOne.Lambda(x, x)) _
        .ToList
    orderingElapsed = Timer - started
    If orderingElapsed < 0 Then orderingElapsed = orderingElapsed + 86400#

    Set dictionary = ROneCOne.DictionaryOf(vbLong, vbLong)
    capacity = dictionary.EnsureCapacity(10000&)
    started = Timer
    For index = 1 To 10000
        dictionary.Add index, index
    Next index
    For index = 1 To 10000
        lastValue = dictionary.Item(index)
    Next index
    hashElapsed10K = Timer - started
    If hashElapsed10K < 0 Then hashElapsed10K = hashElapsed10K + 86400#

    Set dictionary = ROneCOne.DictionaryOf(vbLong, vbLong)
    capacity = dictionary.EnsureCapacity(10000&)
    For index = 1 To 10000
        dictionary.Add index, index
    Next index
    started = Timer
    For index = 1 To 100000
        lastValue = dictionary.Item(((index - 1) Mod 10000) + 1)
    Next index
    hashElapsed100K = Timer - started
    If hashElapsed100K < 0 Then hashElapsed100K = hashElapsed100K + 86400#

    ' Mutation scenario: 10,000 in-place value updates on existing keys plus
    ' 2,000 keyed removals from the tail. Removing from the tail keeps the
    ' array-shift cost near zero, so this isolates the per-operation cost of
    ' maintaining the hash index during mutation.
    Set dictionary = ROneCOne.DictionaryOf(vbLong, vbLong)
    capacity = dictionary.EnsureCapacity(10000&)
    For index = 1 To 10000
        dictionary.Add index, index
    Next index
    started = Timer
    For index = 1 To 10000
        dictionary.Item(index) = dictionary.Item(index) + 1&
    Next index
    For index = 10000 To 8001 Step -1
        dictionary.Remove index
    Next index
    mutationElapsed = Timer - started
    If mutationElapsed < 0 Then mutationElapsed = mutationElapsed + 86400#

    ' List write scenario: 10,000 indexed assignments over a 10,000-element
    ' list. Each write must touch only its array slot, never rebuild the
    ' For Each mirror.
    Set values = ROneCOne.ListOf(vbLong)
    For index = 1 To 10000
        values.Add index
    Next index
    started = Timer
    For index = 0 To 9999
        values.Item(index) = index * 2&
    Next index
    listWriteElapsed = Timer - started
    If listWriteElapsed < 0 Then listWriteElapsed = listWriteElapsed + 86400#
    listWriteCheck = values.Item(9999)

    ' Rows loop scenario: 2,000 positional row reads through table.Rows. The
    ' snapshot list must be version-cached, not rebuilt on every access.
    Set table = ROneCOne.DataTable("BenchmarkRows")
    table.Column "Id", vbLong
    For index = 1 To 2000
        table.LoadRow Array(index)
    Next index
    started = Timer
    rowsTotal = 0
    For index = 0 To 1999
        rowsTotal = rowsTotal + CLng(table.Rows.Item(index).Item("Id"))
    Next index
    rowsLoopElapsed = Timer - started
    If rowsLoopElapsed < 0 Then rowsLoopElapsed = rowsLoopElapsed + 86400#

    ' Generic positional-read scenario: 10,000 indexed reads over a hash set.
    ' A materialized collection must index its own array directly.
    Set hashSet = ROneCOne.HashSetOf(vbLong)
    For index = 1 To 10000
        hashSet.Add index
    Next index
    started = Timer
    For index = 0 To 9999
        genericReadCheck = hashSet.Item(index)
    Next index
    genericReadElapsed = Timer - started
    If genericReadElapsed < 0 Then genericReadElapsed = genericReadElapsed + 86400#

    With ThisWorkbook.Worksheets("Collection Benchmarks")
        .Range("B2").Value2 = 10000
        .Range("B3").Value2 = elapsed
        .Range("B4").Value2 = filtered.Count
        .Range("B5").Value2 = 10000
        .Range("B6").Value2 = orderingElapsed
        .Range("B7").Value2 = ordered.Count
        .Range("B8").Value2 = ordered.First
        .Range("B9").Value2 = 10000
        .Range("B10").Value2 = hashElapsed10K
        .Range("B11").Value2 = 100000
        .Range("B12").Value2 = hashElapsed100K
        .Range("B13").Value2 = lastValue
        .Range("B14").Value2 = 12000
        .Range("B15").Value2 = mutationElapsed
        .Range("B16").Value2 = dictionary.Count
        .Range("B17").Value2 = 10000
        .Range("B18").Value2 = listWriteElapsed
        .Range("B19").Value2 = listWriteCheck
        .Range("B20").Value2 = 2000
        .Range("B21").Value2 = rowsLoopElapsed
        .Range("B22").Value2 = rowsTotal
        .Range("B23").Value2 = 10000
        .Range("B24").Value2 = genericReadElapsed
        .Range("B25").Value2 = genericReadCheck
    End With
End Sub

Private Sub TestKeyedMutationConsistency()
    Dim dictionary As ROneCOne
    Dim enumerated As Long
    Dim index As Long
    Dim removals As Long
    Dim removedValue As Long
    Dim value As Variant
    Dim valueReference As ROneCOne

    ' In-place updates touch only the changed slot, so every other key must
    ' read back its original value and the count must not move.
    Set dictionary = ROneCOne.DictionaryOf(vbLong, vbString)
    For index = 1 To 100
        dictionary.Add index, "seed" & CStr(index)
    Next index
    For index = 1 To 99 Step 2
        dictionary.Item(index) = "updated" & CStr(index)
    Next index
    AssertEqual "mutation updated value", "updated51", dictionary.Item(51&)
    AssertEqual "mutation untouched value", "seed52", dictionary.Item(52&)
    AssertEqual "mutation count after updates", 100&, dictionary.Count

    ' Removals defer the index rebuild; the next read must still resolve
    ' every survivor and reject every removed key.
    removals = 0
    For index = 2 To 100 Step 2
        If dictionary.Remove(index) Then removals = removals + 1
    Next index
    AssertEqual "mutation removal count", 50&, removals
    AssertEqual "mutation count after removals", 50&, dictionary.Count
    AssertFalse "mutation removed key absent", dictionary.ContainsKey(2&)
    AssertTrue "mutation survivor present", dictionary.ContainsKey(3&)
    AssertEqual "mutation survivor value", "updated3", dictionary.Item(3&)

    ' An add after removals inserts through the refreshed index.
    dictionary.Add 2&, "readded2"
    AssertEqual "mutation re-added value", "readded2", dictionary.Item(2&)
    AssertEqual "mutation count after re-add", 51&, dictionary.Count

    ' Enumeration re-materializes from the live arrays after mutation.
    enumerated = 0
    For Each value In dictionary
        enumerated = enumerated + 1
    Next value
    AssertEqual "mutation enumerated count", 51&, enumerated

    ' The concurrent surface reads slot data directly after a probe, so it
    ' must observe a current index across TryUpdate and TryRemove as well.
    Set dictionary = ROneCOne.ConcurrentDictionaryOf(vbLong, vbLong)
    For index = 1 To 50
        dictionary.Add index, index
    Next index
    AssertTrue "mutation TryUpdate", dictionary.TryUpdate(7&, 70&, 7&)
    AssertEqual "mutation TryUpdate value", 70&, dictionary.Item(7&)
    Set valueReference = ROneCOne.RefLong(removedValue)
    AssertTrue "mutation TryRemove", dictionary.TryRemove(8&, valueReference)
    AssertEqual "mutation TryRemove value", 8&, removedValue
    AssertFalse "mutation TryRemove key absent", dictionary.ContainsKey(8&)
    AssertEqual "mutation AddOrUpdate", 700&, dictionary.AddOrUpdate(7&, 1&, 700&)

    ' Replacing a hash set element by index swaps the hash key itself, so
    ' membership must follow the new element on the very next read.
    Set dictionary = ROneCOne.HashSetOf(vbLong)
    dictionary.Add 1&
    dictionary.Add 2&
    dictionary.Item(0) = 9&
    AssertTrue "mutation set replacement present", dictionary.Contains(9&)
    AssertFalse "mutation set replacement removed", dictionary.Contains(1&)
End Sub

Private Sub TestListMirrorConsistency()
    Dim joined As String
    Dim total As Long
    Dim value As Variant
    Dim values As ROneCOne

    ' The For Each mirror is rebuilt lazily from the element array, so every
    ' structural mutation must be visible to the next enumeration.
    Set values = ROneCOne.ListOf(vbLong, CLng(1), CLng(2), CLng(3))
    total = 0
    For Each value In values
        total = total + CLng(value)
    Next value
    AssertEqual "mirror initial enumeration", 6&, total

    ' An indexed write must show its new value on the next For Each.
    values.Item(1) = 20&
    joined = vbNullString
    For Each value In values
        joined = joined & CStr(value) & ";"
    Next value
    AssertEqual "mirror after indexed write", "1;20;3;", joined

    ' Insert, RemoveAt, and Remove keep enumeration order and membership.
    values.Insert 0, 9&
    joined = vbNullString
    For Each value In values
        joined = joined & CStr(value) & ";"
    Next value
    AssertEqual "mirror after insert", "9;1;20;3;", joined
    values.RemoveAt 0
    values.Remove 20&
    joined = vbNullString
    For Each value In values
        joined = joined & CStr(value) & ";"
    Next value
    AssertEqual "mirror after removals", "1;3;", joined

    ' AddRange appends atomically; Clear empties the next enumeration.
    values.AddRange Array(CLng(7), CLng(8))
    total = 0
    For Each value In values
        total = total + CLng(value)
    Next value
    AssertEqual "mirror after AddRange", 19&, total
    values.Clear
    total = 0
    For Each value In values
        total = total + 1
    Next value
    AssertEqual "mirror after Clear", 0&, total
    values.Add 5&
    total = 0
    For Each value In values
        total = total + CLng(value)
    Next value
    AssertEqual "mirror after re-add", 5&, total
    AssertEqual "list count fast path", 1&, values.Count
End Sub

Private Sub TestCompleteLinqSurface()
    Dim addValues As ROneCOne
    Dim chunks As ROneCOne
    Dim dictionary As ROneCOne
    Dim firstChunk As ROneCOne
    Dim firstList As ROneCOne
    Dim flattened As ROneCOne
    Dim group As ROneCOne
    Dim groups As ROneCOne
    Dim joined As ROneCOne
    Dim lookup As ROneCOne
    Dim nested As ROneCOne
    Dim numbers As ROneCOne
    Dim other As ROneCOne
    Dim secondList As ROneCOne
    Dim strings As ROneCOne
    Dim x As ROneCOne
    Dim y As ROneCOne

    Set numbers = ROneCOne.ListOf( _
        vbLong, CLng(1), CLng(2), CLng(3), CLng(4), CLng(5))
    Set other = ROneCOne.ListOf(vbLong, CLng(3), CLng(4), CLng(6))
    Set x = ROneCOne.Var(vbLong)
    Set y = ROneCOne.Var(vbLong)
    Set addValues = x.Add(y).AsFunc

    mCurrentTest = "CompleteLinq TakeWhile"
    AssertEqual "TakeWhile", CLng(3), _
        numbers.TakeWhile(numbers.Element.LessThan(CLng(4))).Count
    mCurrentTest = "CompleteLinq SkipWhile"
    AssertEqual "SkipWhile", CLng(2), _
        numbers.SkipWhile(numbers.Element.LessThan(CLng(4))).Count
    mCurrentTest = "CompleteLinq TakeLast"
    AssertEqual "TakeLast", CLng(4), numbers.TakeLast(2).Item(0)
    mCurrentTest = "CompleteLinq SkipLast"
    AssertEqual "SkipLast", CLng(3), numbers.SkipLast(2).Count
    mCurrentTest = "CompleteLinq Concat"
    AssertEqual "Concat", CLng(8), numbers.Concat(other).Count
    mCurrentTest = "CompleteLinq Union"
    AssertEqual "Union", CLng(6), numbers.Union(other).Count
    mCurrentTest = "CompleteLinq Intersect"
    AssertEqual "Intersect", CLng(2), numbers.Intersect(other).Count
    mCurrentTest = "CompleteLinq Except"
    AssertEqual "Except", CLng(3), numbers.Except(other).Count
    mCurrentTest = "CompleteLinq DefaultIfEmpty"
    AssertEqual "DefaultIfEmpty", CLng(0), _
        ROneCOne.ListOf(vbLong).DefaultIfEmpty.Item(0)
    mCurrentTest = "CompleteLinq ElementAt"
    AssertEqual "ElementAt", CLng(3), numbers.ElementAt(2)
    mCurrentTest = "CompleteLinq ElementAtOrDefault"
    AssertEqual "ElementAtOrDefault", CLng(0), numbers.ElementAtOrDefault(99)
    mCurrentTest = "CompleteLinq Aggregate"
    AssertEqual "Aggregate", CLng(15), numbers.Aggregate(CLng(0), addValues)

    mCurrentTest = "CompleteLinq chunks"
    Set chunks = numbers.Chunk(2).ToList
    Set firstChunk = chunks.Item(0)
    AssertEqual "Chunk count", CLng(3), chunks.Count
    AssertEqual "Chunk element count", CLng(2), firstChunk.Count

    mCurrentTest = "CompleteLinq SelectMany"
    Set firstList = ROneCOne.ListOf(vbLong, CLng(1), CLng(2))
    Set secondList = ROneCOne.ListOf(vbLong, CLng(3), CLng(4))
    Set nested = ROneCOne.ListOf(firstList, firstList, secondList)
    Set flattened = nested.SelectMany(nested.Element, vbLong).ToList
    AssertEqual "SelectMany", CLng(4), flattened.Count

    mCurrentTest = "CompleteLinq keyed materializers"
    Set dictionary = numbers.ToDictionary(numbers.Element)
    AssertEqual "ToDictionary", CLng(4), dictionary(CLng(4))
    AssertEqual "ToHashSet", CLng(5), numbers.ToHashSet.Count

    mCurrentTest = "CompleteLinq lookup"
    Set strings = ROneCOne.ListOf( _
        vbString, "ada", "alan", "grace", "guido")
    Set lookup = strings.ToLookup( _
        ROneCOne.Func("DelegateProcedures.FirstCharacter"))
    Set group = lookup.Item("a")
    AssertEqual "ToLookup", CLng(2), group.Count

    mCurrentTest = "CompleteLinq GroupBy"
    Set groups = strings.GroupBy( _
        ROneCOne.Func("DelegateProcedures.FirstCharacter"))
    Set group = groups.Item(0)
    AssertEqual "GroupBy key", "a", group.Key
    AssertEqual "GroupBy group count", CLng(2), group.Count

    mCurrentTest = "CompleteLinq Join and Zip"
    Set joined = numbers.Join(other, numbers.Element, other.Element, addValues)
    AssertEqual "Join count", CLng(2), joined.Count
    AssertEqual "Join result", CLng(6), joined.Item(0)
    AssertEqual "Zip", CLng(4), _
        numbers.Zip(other, addValues).Item(0)
End Sub

Private Sub TestModernLinqSurface()
    Dim addValues As ROneCOne
    Dim counts As ROneCOne
    Dim indexed As ROneCOne
    Dim mixed As ROneCOne
    Dim numbers As ROneCOne
    Dim other As ROneCOne
    Dim parity As ROneCOne
    Dim reflectedCount As Long
    Dim reflectedCountOut As ROneCOne
    Dim sums As ROneCOne
    Dim x As ROneCOne
    Dim y As ROneCOne

    Set numbers = ROneCOne.ListOf( _
        vbLong, 1&, 2&, 3&, 4&, 5&)
    Set other = ROneCOne.ListOf(vbLong, 3&, 4&, 6&)
    Set x = ROneCOne.Var(vbLong)
    Set y = ROneCOne.Var(vbLong)
    Set parity = x.Modulo(2&).AsFunc
    Set addValues = x.Add(y).AsFunc

    AssertEqual "Empty", 0&, ROneCOne.EmptyOf(vbLong).Count
    AssertTrue "AsEnumerable identity", numbers.AsEnumerable Is numbers
    AssertEqual "LongCount", CLngLng(5), numbers.LongCount
    Set reflectedCountOut = ROneCOne.RefLong(reflectedCount)
    AssertTrue "TryGetNonEnumeratedCount list", _
        numbers.TryGetNonEnumeratedCount(reflectedCountOut)
    AssertEqual "TryGetNonEnumeratedCount value", 5&, reflectedCount
    AssertFalse "TryGetNonEnumeratedCount query", _
        numbers.Where(numbers.Element.AtLeast(1&)) _
            .TryGetNonEnumeratedCount(reflectedCountOut)

    Set mixed = ROneCOne.ListOf(vbVariant, 1&, "two", 3&)
    AssertEqual "OfType", 2&, mixed.OfType(vbLong).Count
    AssertEqual "Cast", 2&, _
        ROneCOne.ListOf(vbVariant, 1&, 2&).Cast(vbLong).Item(1)

    AssertEqual "UnionBy", 4&, _
        numbers.UnionBy(other, x.Modulo(4&).AsFunc).Count
    AssertEqual "IntersectBy", 1&, _
        numbers.IntersectBy(ROneCOne.ListOf(vbLong, 0&), parity).Count
    AssertEqual "ExceptBy", 1&, _
        numbers.ExceptBy(ROneCOne.ListOf(vbLong, 0&), parity).Count

    Set counts = numbers.CountBy(parity)
    AssertEqual "CountBy odd", 3&, counts(1&)
    AssertEqual "CountBy even", 2&, counts(0&)
    Set sums = numbers.AggregateBy(parity, 0&, addValues)
    AssertEqual "AggregateBy odd", 9&, sums(1&)
    AssertEqual "AggregateBy even", 6&, sums(0&)

    Set indexed = numbers.Index
    AssertEqual "Index key", 1&, indexed.Item(1).Key
    AssertEqual "Index value", 2&, indexed.Item(1).Value
    AssertEqual "ToDictionary element selector", 30&, _
        numbers.ToDictionary(x, x.Multiply(10&).AsFunc)(3&)
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

Private Sub TestNumericWidening()
    Dim discount As ROneCOne
    Dim doubles As ROneCOne
    Dim longs As ROneCOne
    Dim price As ROneCOne
    Dim scores As ROneCOne
    Dim singles As ROneCOne
    Dim typedFunc As ROneCOne
    Dim actualError As Long

    ' Narrower integer literals widen into a wider integer list, and each stored
    ' value takes the declared element type exactly.
    Set longs = ROneCOne.ListOf(vbLong, CInt(1), CByte(2), CLng(3))
    AssertEqual "widen count", CLng(3), longs.Count
    AssertEqual "widened value", CLng(1), longs.Item(0)
    AssertEqual "widened stored type", "Long", TypeName(longs.Item(0))
    AssertEqual "widen keeps generic name", "List<Long>", longs.GenericTypeName

    ' Integers, Long, and Single all widen losslessly into a Double list.
    Set doubles = ROneCOne.ListOf(vbDouble)
    doubles.Add CLng(100000)
    doubles.Add CInt(7)
    doubles.Add CSng(1.5)
    AssertEqual "widened to double type", "Double", TypeName(doubles.Item(0))
    AssertEqual "widened double value", CDbl(7), doubles.Item(1)

    ' Narrowing (Double into Long) is still rejected, atomically.
    Set longs = ROneCOne.ListOf(vbLong)
    On Error Resume Next
    longs.Add CDbl(1.5)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "reject narrowing double to long", _
        ROneCOne.TypeMismatchError, actualError
    AssertEqual "rejected add is atomic", CLng(0), longs.Count

    ' Long into Single is rejected as a class because some Longs lose precision;
    ' Integer into Single is accepted because it never does.
    Set singles = ROneCOne.ListOf(vbSingle)
    On Error Resume Next
    singles.Add CLng(100000)
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "reject lossy long to single", _
        ROneCOne.TypeMismatchError, actualError
    singles.Add CInt(300)
    AssertEqual "accept integer to single", CLng(1), singles.Count

    ' Dictionary keys widen for both insertion and indexed lookup.
    Set scores = ROneCOne.DictionaryOf(vbLong, vbLong)
    scores.Add CInt(101), CInt(95)
    AssertEqual "widened dictionary lookup", CLng(95), scores.Item(CInt(101))

    ' A Double-typed lambda accepts an integer argument with no CDbl ceremony.
    Set price = ROneCOne.Var(vbDouble)
    Set discount = price.Multiply(0.9).AsFunc
    AssertEqual "widened lambda argument", CDbl(90), discount(CInt(100))

    ' A Func declared to return Double accepts a Long-producing body.
    Set typedFunc = ROneCOne.Value(CLng(42)).AsFunc.Returns(vbDouble)
    AssertEqual "widened func result type", "Double", TypeName(typedFunc.Run())
End Sub

Private Sub TestExpressionDisplay()
    Dim customers As ROneCOne
    Dim price As ROneCOne
    Dim rule As ROneCOne

    ' A reusable pricing rule renders as the arithmetic a reader would expect.
    Set price = ROneCOne.Var(vbDouble)
    Set rule = price.Multiply(0.9).AsFunc
    AssertEqual "lambda display", "(x * 0.9)", rule.ToDisplayString

    ' A deferred member condition prints as readable pseudocode.
    Set customers = ROneCOne.ListOf(vbObject)
    AssertEqual "member condition display", "(x.Age >= 40)", _
        customers.Condition("Age").AtLeast(40).ToDisplayString
End Sub

Private Sub TestIndexedAccessScales()
    Dim expected As Double
    Dim i As Long
    Dim numbers As ROneCOne
    Dim total As Double

    Set numbers = ROneCOne.ListOf(vbLong)
    For i = 1 To 10000
        numbers.Add i
    Next i
    ' Sum through the numeric indexer. With array-backed storage each access is
    ' O(1), so this whole pass is linear; under a linked Collection it would be
    ' quadratic and blow past the harness deadline on ten thousand elements.
    For i = 0 To numbers.Count - 1
        total = total + numbers(i)
    Next i
    expected = 10000# * 10001# / 2#
    AssertEqual "indexed access sum", expected, total
    AssertEqual "indexed access count", 10000&, numbers.Count
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

Private Sub TestOrderingSystem()
    Dim actualError As Long
    Dim ada As GenericCustomer
    Dim booleans As ROneCOne
    Dim comparer As ROneCOne
    Dim customers As ROneCOne
    Dim grace As GenericCustomer
    Dim katherine As GenericCustomer
    Dim margaret As GenericCustomer
    Dim mixedValues As ROneCOne
    Dim ordered As ROneCOne
    Dim prototype As GenericCustomer
    Dim query As ROneCOne
    Dim selector As ROneCOne
    Dim strings As ROneCOne
    Dim values As ROneCOne
    Dim variantValues As ROneCOne

    Set values = ROneCOne.ListOf(vbLong, CLng(3), CLng(1), CLng(2))
    Set ordered = values.Order.ToList
    AssertEqual "Order first", CLng(1), ordered.First
    AssertEqual "Order last", CLng(3), ordered.Last
    Set ordered = values.OrderDescending.ToList
    AssertEqual "OrderDescending first", CLng(3), ordered.First
    AssertEqual "OrderDescending last", CLng(1), ordered.Last

    Set strings = ROneCOne.ListOf(vbString, "a", "Z")
    AssertEqual "ordinal string order", "Z", strings.Order.First
    Set booleans = ROneCOne.ListOf(vbBoolean, True, False)
    AssertTrue "Boolean False first", Not CBool(booleans.Order.First)

    Set variantValues = ROneCOne.ListOf(vbVariant, "b", Null, "a")
    Set ordered = variantValues.Order.ToList
    AssertTrue "Null first ascending", IsNull(ordered.First)
    Set ordered = variantValues.OrderDescending.ToList
    AssertTrue "Null last descending", IsNull(ordered.Last)

    Set mixedValues = ROneCOne.ListOf(vbVariant, CLng(1), "2")
    On Error Resume Next
    Set ordered = mixedValues.Order.ToList
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "mixed Variant keys reject coercion", _
        ROneCOne.TypeMismatchError, actualError

    Set prototype = New GenericCustomer
    Set ada = New GenericCustomer
    ada.CustomerName = "Ada"
    ada.Age = 36
    ada.City = "London"
    Set grace = New GenericCustomer
    grace.CustomerName = "Grace"
    grace.Age = 40
    grace.City = "London"
    Set katherine = New GenericCustomer
    katherine.CustomerName = "Katherine"
    katherine.Age = 49
    katherine.City = "Cleveland"
    Set margaret = New GenericCustomer
    margaret.CustomerName = "Margaret"
    margaret.Age = 40
    margaret.City = "London"
    Set customers = ROneCOne.ListOf(prototype, ada, grace, katherine, margaret)

    Set ordered = customers _
        .OrderBy("City") _
        .ThenByDescending("Age") _
        .ToList
    AssertEqual "composite primary", "Katherine", ordered.Item(0).CustomerName
    AssertEqual "composite secondary", "Grace", ordered.Item(1).CustomerName
    AssertEqual "stable equal keys", "Margaret", ordered.Item(2).CustomerName
    AssertEqual "composite tail", "Ada", ordered.Item(3).CustomerName

    Set ordered = customers _
        .OrderBy("City") _
        .ThenBy("Age") _
        .ThenBy("CustomerName") _
        .ToList
    AssertEqual "ThenBy ascending", "Ada", ordered.Item(1).CustomerName
    AssertEqual "multiple ThenBy", "Grace", ordered.Item(2).CustomerName
    Set ordered = customers _
        .OrderByDescending("Age") _
        .ThenBy("CustomerName") _
        .ToList
    AssertEqual "OrderByDescending", "Katherine", ordered.First.CustomerName

    Set comparer = ROneCOne.Comparer( _
        "DelegateProcedures.CompareTextIgnoreCase")
    Set ordered = customers _
        .OrderBy("City", comparer) _
        .ThenBy("CustomerName", comparer) _
        .ToList
    AssertEqual "per-level comparer", "Katherine", ordered.First.CustomerName

    Set comparer = ROneCOne.Comparer( _
        "DelegateProcedures.CompareCustomerAge")
    Set ordered = customers.Order(comparer).ToList
    AssertEqual "object comparer", "Ada", ordered.First.CustomerName
    On Error Resume Next
    Set ordered = customers.Order.ToList
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "objects require comparer", ROneCOne.TypeMismatchError, actualError

    On Error Resume Next
    Set ordered = values.ThenBy(values.Element).ToList
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "ThenBy requires ordered query", _
        ROneCOne.InvalidOperationError, actualError
    Set query = values.Where(values.Element.AtLeast(CLng(1)))
    On Error Resume Next
    Set ordered = query.ThenBy(query.Element).ToList
    actualError = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual "Where terminates ordered query", _
        ROneCOne.InvalidOperationError, actualError

    Set ordered = customers _
        .OrderByDescending("Age") _
        .OrderBy("CustomerName") _
        .ToList
    AssertEqual "OrderBy resets chain", "Ada", ordered.First.CustomerName

    Set query = values.Order
    values.Add CLng(0)
    Set ordered = query.ToList
    AssertEqual "deferred Order sees mutation", CLng(0), ordered.First

    Set selector = ROneCOne.Func( _
        "DelegateProcedures.CountedLongIdentity") _
        .Takes(vbLong) _
        .Returns(vbLong)
    DelegateProcedures.ResetSelectorCalls
    Set ordered = values _
        .OrderBy(selector) _
        .ThenBy(selector) _
        .ToList
    AssertEqual "selectors evaluate once per level", _
        values.Count * CLng(2), DelegateProcedures.CurrentSelectorCalls
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

Private Sub AssertEqual( _
    ByVal testName As String, _
    ByVal expected As Variant, _
    ByVal actual As Variant _
)
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

Private Sub AssertFalse(ByVal testName As String, ByVal condition As Boolean)
    AssertTrue testName, Not condition
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
