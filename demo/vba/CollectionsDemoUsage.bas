Attribute VB_Name = "CollectionsDemoUsage"
Option Explicit

' This module is intentionally organized as a small executable tutorial.
' The public macro coordinates the demo; each private procedure owns one topic.

Private Const BASIC_EXAMPLES_SHEET As String = "Examples"
Private Const BENCHMARK_ELEMENT_COUNT As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const START_SHEET As String = "Start Here"
Private Const USER_CLASS_SHEET As String = "User Class LINQ"

Private mForEachTotal As Long

' -----------------------------------------------------------------------------
' Entry point
' -----------------------------------------------------------------------------

Public Sub RunROneCOneCollectionsDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure

    WritePrimitiveCollectionExamples
    WriteUserClassLinqExamples
    RunCollectionBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

' -----------------------------------------------------------------------------
' Primitive List<T> and scalar LINQ
' -----------------------------------------------------------------------------

Private Sub WritePrimitiveCollectionExamples()
    Dim ada As DemoCustomer
    Dim customers As ROneCOne
    Dim enumerationValues As ROneCOne
    Dim filtered As ROneCOne
    Dim grace As DemoCustomer
    Dim numbers As ROneCOne
    Dim projected As ROneCOne
    Dim sequence As ROneCOne
    Dim strictTypeRejected As Boolean
    Dim total As Long
    Dim x As ROneCOne

    Set numbers = ROneCOne.ListOf( _
        vbLong, CLng(5), CLng(10), CLng(15))
    strictTypeRejected = RaisesExpectedTypeMismatch(numbers)

    ' ListFrom infers the exact user class from its first value.
    Set ada = CreateCustomer("Ada", CLng(36), "London")
    Set grace = CreateCustomer("Grace", CLng(40), "Arlington")
    Set customers = ROneCOne.ListFrom(ada, grace)

    ' Deferred execution means this query observes the later Add operation.
    Set numbers = ROneCOne.ListOf(vbLong, CLng(5), CLng(20))
    Set x = numbers.Element
    Set filtered = numbers.Where(x.GreaterThan(CLng(10)))
    numbers.Add CLng(30)
    Set filtered = filtered.ToList

    Set projected = ROneCOne.Range(CLng(1), CLng(6)) _
        .Where(x.Modulo(CLng(2)).EqualTo(CLng(0))) _
        .Map(x.Multiply(CLng(10)), vbLong) _
        .OrderDescending _
        .Take(CLng(2)) _
        .ToList

    Set sequence = ROneCOne.ListOf(vbLong, CLng(2), CLng(2), CLng(3))
    Set sequence = sequence.Distinct _
        .Prepend(CLng(1)) _
        .Append(CLng(4)) _
        .Reverse _
        .Skip(CLng(1)) _
        .ToList

    Set numbers = ROneCOne.Range(CLng(1), CLng(5))
    Set enumerationValues = ROneCOne.Range(CLng(1), CLng(4))
    mForEachTotal = 0
    enumerationValues.ForEach ROneCOne.Action( _
        "CollectionsDemoUsage.DemoAccumulateLong").Takes(vbLong)
    total = mForEachTotal

    With ThisWorkbook.Worksheets(BASIC_EXAMPLES_SHEET)
        .Range("E6").Value2 = numbers.GenericTypeName
        .Range("E7").Value2 = strictTypeRejected
        .Range("E8").Value2 = customers.GenericTypeName & ":" & _
            customers.Item(1).CustomerName
        .Range("E9").Value2 = CStr(filtered.Count) & "|" & CStr(filtered.Last)
        .Range("E10").Value2 = projected.JoinText(",")
        .Range("E11").Value2 = sequence.JoinText(",")
        .Range("E12").Value2 = CStr(numbers.Sum) & "|" & _
            CStr(numbers.Average) & "|" & CStr(numbers.Min) & "|" & _
            CStr(numbers.Max)
        .Range("E13").Value2 = total
    End With
End Sub

' -----------------------------------------------------------------------------
' User-defined class LINQ
' -----------------------------------------------------------------------------

Private Sub WriteUserClassLinqExamples()
    Dim Age As Variant
    Dim allExperienced As Boolean
    Dim anyLondon As Boolean
    Dim ada As DemoCustomer
    Dim allowedCities As ROneCOne
    Dim comparer As ROneCOne
    Dim City As Variant
    Dim customers As ROneCOne
    Dim distinctCities As ROneCOne
    Dim experienced As ROneCOne
    Dim equalityComparer As ROneCOne
    Dim firstCustomer As DemoCustomer
    Dim grace As DemoCustomer
    Dim katherine As DemoCustomer
    Dim lastCustomer As DemoCustomer
    Dim managed As ROneCOne
    Dim margaret As DemoCustomer
    Dim membershipMatches As ROneCOne
    Dim names As ROneCOne
    Dim orderedCustomers As ROneCOne
    Dim procedureFiltered As ROneCOne
    Dim reportPredicate As ROneCOne
    Dim singleCustomer As DemoCustomer
    Dim strings As ROneCOne
    Dim stringMatches As ROneCOne

    Set ada = CreateCustomer("Ada", CLng(36), "London")
    Set grace = CreateCustomer("Grace", CLng(40), "Arlington")
    Set grace.Manager = ada
    Set katherine = CreateCustomer("Katherine", CLng(49), "Cleveland")
    Set katherine.Manager = grace
    Set customers = ROneCOne.ListFrom(ada, grace, katherine)

    ' Build the query first, then mutate its source to prove it remains deferred.
    Set experienced = customers.Where("Age").AtLeast(CLng(40))
    Set margaret = CreateCustomer("Margaret", CLng(45), "Arlington")
    Set margaret.Manager = grace
    customers.Add margaret
    Set lastCustomer = experienced.Last

    ' Member-name selectors eliminate the explicit element parameter.
    Set names = experienced _
        .Map("CustomerName", vbString) _
        .Order _
        .ToList

    ' Each ThenBy level stays typed, deferred, stable, and independently directed.
    Set orderedCustomers = customers _
        .OrderBy("City") _
        .ThenByDescending("Age") _
        .ToList
    Set firstCustomer = orderedCustomers.First

    ' Condition creates a reusable expression when a predicate is composed.
    anyLondon = customers.Exists( _
        customers.Condition("City").EqualTo("London"))
    allExperienced = customers.All( _
        customers.Condition("Age").AtLeast(CLng(40)))

    ' Procedure signatures are inferred as Func<DemoCustomer, Boolean>.
    Set procedureFiltered = customers.WhereMethod( _
        "CollectionsDemoUsage.IsExperiencedCustomer").ToList

    Set stringMatches = customers _
        .Where("CustomerName").StartsWith("G") _
        .ToList
    Set distinctCities = customers.DistinctBy("City").ToList

    ' C#-style ?. safely propagates Null when an intermediate object is Nothing.
    Set managed = customers.Where("Manager?.Age").AtLeast(CLng(40)).ToList

    Set allowedCities = ROneCOne.ListOf( _
        vbString, "London", "Cleveland")
    Set equalityComparer = ROneCOne.EqualityComparer( _
        "CollectionsDemoUsage.DemoTextEqualsIgnoreCase")
    Set comparer = ROneCOne.Comparer( _
        "CollectionsDemoUsage.DemoCompareTextIgnoreCase")
    Set strings = ROneCOne.ListOf( _
        vbString, "Ada", "ADA", "grace")
    Set singleCustomer = customers.SingleItem( _
        customers.Match("CustomerName", "Grace"))

    Set ada.Reports = ROneCOne.ListFrom(grace, katherine)
    Set grace.Reports = ROneCOne.ListFrom(ada)
    Set katherine.Reports = ROneCOne.ListLike(ada)
    Set margaret.Reports = ROneCOne.ListFrom(grace)
    Set reportPredicate = ada.Reports.Condition("Age").AtLeast(CLng(40))
    Set membershipMatches = customers.Where(allowedCities.Contains(customers!City))

    With ThisWorkbook.Worksheets(USER_CLASS_SHEET)
        .Range("E6").Value2 = customers.GenericTypeName & ":" & customers.Count
        .Range("E7").Value2 = CStr(experienced.Count) & "|" & _
            lastCustomer.CustomerName
        .Range("E8").Value2 = names.JoinText("|")
        .Range("E9").Value2 = firstCustomer.CustomerName & "|" & _
            CStr(firstCustomer.Age)
        .Range("E10").Value2 = CStr(anyLondon) & "|" & CStr(allExperienced)
        .Range("E11").Value2 = Round(CDbl( _
            experienced.Average("Age")), 1)
        .Range("E12").Value2 = procedureFiltered.Count & "|" & _
            customers.Predicate( _
                "CollectionsDemoUsage.IsExperiencedCustomer").Signature
        .Range("E13").Value2 = stringMatches.Count & "|" & _
            customers.Where("CustomerName").Contains("ther").Count
        .Range("E14").Value2 = distinctCities.Count
        With customers
            Set names = .Where(!Age.AtLeast(CLng(40))) _
                .Map("CustomerName", vbString) _
                .Order _
                .ToList
        End With
        .Range("E15").Value2 = names.JoinText("|")
        .Range("E16").Value2 = customers.Where("Manager").IsNothing.Count & _
            "|" & managed.Count
        .Range("E17").Value2 = customers.Where("City").IsIn( _
            allowedCities).Count & "|" & _
            membershipMatches.Count
        .Range("E18").Value2 = customers.Where("City") _
            .EqualToIgnoreCase("LONDON").Count & "|" & _
            customers.Where("CustomerName") _
                .ContainsIgnoreCase("THER").Count
        .Range("E19").Value2 = customers.Count( _
            customers.Condition("Age").AtLeast(CLng(40))) & "|" & _
            singleCustomer.CustomerName & "|" & _
            customers.None(customers.Match("Age", CLng(100)))
        .Range("E20").Value2 = customers.WhereAny( _
            "Reports", reportPredicate).Count & "|" & _
            customers.WhereAll("Reports", ada.Reports.Condition("Age") _
                .AtLeast(CLng(36))).Count & "|" & _
            customers.WhereNone("Reports", reportPredicate).Count
        .Range("E21").Value2 = strings.Distinct( _
            equalityComparer).Count & "|" & _
            strings.Contains("ada", equalityComparer) & "|" & _
            strings.Order(comparer).First
        .Range("E22").Value2 = customers.Where( _
            customers.Condition("Age").AtLeast(CLng(40)).Both( _
                customers.Match("City", "Arlington"))).Count & "|" & _
            customers.Where(customers.Match("City", "London").Either( _
                customers.Match("City", "Cleveland"))).Count
    End With
End Sub

' -----------------------------------------------------------------------------
' Benchmark and reporting
' -----------------------------------------------------------------------------

Private Sub RunCollectionBenchmark()
    Dim capacity As Long
    Dim customer As DemoCustomer
    Dim customers As ROneCOne
    Dim dictionary As ROneCOne
    Dim filtered As ROneCOne
    Dim index As Long
    Dim lastValue As Long
    Dim numbers As ROneCOne
    Dim ordered As ROneCOne
    Dim started As Double
    Dim x As ROneCOne

    started = Timer
    Set numbers = ROneCOne.Range(CLng(1), BENCHMARK_ELEMENT_COUNT)
    Set x = numbers.Element
    Set filtered = numbers _
        .Where(x.Modulo(CLng(2)).EqualTo(CLng(0))) _
        .ToList

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = filtered.Count
    End With

    Set customer = CreateCustomer("Benchmark", CLng(50), "Local")
    Set customers = ROneCOne.Repeat(customer, BENCHMARK_ELEMENT_COUNT)
    started = Timer
    Set filtered = customers.Where("Age").AtLeast(CLng(40)).ToList
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B7").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C7").Value2 = ElapsedSeconds(started)
        .Range("D7").Value2 = filtered.Count
    End With

    started = Timer
    Set ordered = numbers _
        .OrderBy(x.Modulo(CLng(100))) _
        .ThenByDescending(x) _
        .ToList
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B8").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C8").Value2 = ElapsedSeconds(started)
        .Range("D8").Value2 = ordered.Count
    End With

    Set dictionary = ROneCOne.DictionaryOf(vbLong, vbLong)
    capacity = dictionary.EnsureCapacity(BENCHMARK_ELEMENT_COUNT)
    started = Timer
    For index = 1 To BENCHMARK_ELEMENT_COUNT
        dictionary.Add index, index
    Next index
    For index = 1 To BENCHMARK_ELEMENT_COUNT
        lastValue = dictionary.Item(index)
    Next index
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B9").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C9").Value2 = ElapsedSeconds(started)
        .Range("D9").Value2 = lastValue
    End With

    started = Timer
    For index = 1 To 100000
        lastValue = dictionary.Item(((index - 1) Mod BENCHMARK_ELEMENT_COUNT) + 1)
    Next index
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B10").Value2 = 100000
        .Range("C10").Value2 = ElapsedSeconds(started)
        .Range("D10").Value2 = lastValue
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

' -----------------------------------------------------------------------------
' Small reusable helpers
' -----------------------------------------------------------------------------

Private Function CreateCustomer( _
    ByVal customerName As String, _
    ByVal age As Long, _
    ByVal city As String _
) As DemoCustomer
    Dim customer As DemoCustomer

    Set customer = New DemoCustomer
    customer.CustomerName = customerName
    customer.Age = age
    customer.City = city
    Set CreateCustomer = customer
End Function

Public Sub DemoAccumulateLong(ByVal value As Variant)
    mForEachTotal = mForEachTotal + CLng(value)
End Sub

Public Function IsExperiencedCustomer(ByVal value As Variant) As Variant
    IsExperiencedCustomer = (CLng(value.Age) >= 40)
End Function

Public Function DemoTextEqualsIgnoreCase( _
    ByVal leftValue As Variant, _
    ByVal rightValue As Variant _
) As Variant
    DemoTextEqualsIgnoreCase = (StrComp( _
        CStr(leftValue), CStr(rightValue), vbTextCompare) = 0)
End Function

Public Function DemoCompareTextIgnoreCase( _
    ByVal leftValue As Variant, _
    ByVal rightValue As Variant _
) As Variant
    DemoCompareTextIgnoreCase = CLng(Sgn(StrComp( _
        CStr(leftValue), CStr(rightValue), vbTextCompare)))
End Function

Private Function RaisesExpectedTypeMismatch(ByVal values As ROneCOne) As Boolean
    Dim description As String
    Dim number As Long

    On Error GoTo TypeMismatch
    values.Add "not a Long"
    Exit Function

TypeMismatch:
    number = Err.Number
    description = Err.Description
    Err.Clear
    If number = ROneCOne.TypeMismatchError Then
        RaisesExpectedTypeMismatch = True
    Else
        Err.Raise number, "CollectionsDemoUsage", description
    End If
End Function

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
