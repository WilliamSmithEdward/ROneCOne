Attribute VB_Name = "CollectionsDemoUsage"
Option Explicit

' This module is intentionally organized as a small executable tutorial.
' The public macro coordinates the demo; each private procedure owns one topic.

Private Const BASIC_EXAMPLES_SHEET As String = "Examples"
Private Const BENCHMARK_ELEMENT_COUNT As Long = 10000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const START_SHEET As String = "Start Here"
Private Const USER_CLASS_SHEET As String = "User Class LINQ"

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
    Dim customerPrototype As DemoCustomer
    Dim enumerationValues As ROneCOne
    Dim filtered As ROneCOne
    Dim grace As DemoCustomer
    Dim numbers As ROneCOne
    Dim projected As ROneCOne
    Dim sequence As ROneCOne
    Dim strictTypeRejected As Boolean
    Dim total As Long
    Dim value As Variant
    Dim x As ROneCOne

    Set numbers = ROneCOne.ListOf(vbLong)
    numbers.Add CLng(5)
    numbers.Add CLng(10)
    numbers.Add CLng(15)
    strictTypeRejected = RaisesExpectedTypeMismatch(numbers)

    ' A prototype captures T without being retained by the collection.
    Set customerPrototype = New DemoCustomer
    Set customers = ROneCOne.ListOf(customerPrototype)
    Set customerPrototype = Nothing
    Set ada = CreateCustomer("Ada", CLng(36), "London")
    Set grace = CreateCustomer("Grace", CLng(40), "Arlington")
    customers.Add ada
    customers.Add grace

    ' Deferred execution means this query observes the later Add operation.
    Set numbers = ROneCOne.ListOf(vbLong)
    numbers.Add CLng(5)
    numbers.Add CLng(20)
    Set x = ROneCOne.Parameter(vbLong)
    Set filtered = numbers.Where( _
        ROneCOne.Lambda(x.GreaterThan(CLng(10)), x))
    numbers.Add CLng(30)
    Set filtered = filtered.ToList

    Set projected = ROneCOne.Range(CLng(1), CLng(6)) _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .SelectItems(ROneCOne.Lambda(x.Multiply(CLng(10)), x), vbLong) _
        .OrderByDescending(ROneCOne.Lambda(x, x)) _
        .Take(CLng(2)) _
        .ToList

    Set sequence = ROneCOne.ListOf(vbLong)
    sequence.Add CLng(2)
    sequence.Add CLng(2)
    sequence.Add CLng(3)
    Set sequence = sequence.Distinct _
        .Prepend(CLng(1)) _
        .Append(CLng(4)) _
        .Reverse _
        .Skip(CLng(1)) _
        .ToList

    Set numbers = ROneCOne.Range(CLng(1), CLng(5))
    Set enumerationValues = ROneCOne.Range(CLng(1), CLng(4))
    For Each value In enumerationValues
        total = total + CLng(value)
    Next value

    With ThisWorkbook.Worksheets(BASIC_EXAMPLES_SHEET)
        .Range("E6").Value2 = numbers.GenericTypeName
        .Range("E7").Value2 = strictTypeRejected
        .Range("E8").Value2 = customers.GenericTypeName & ":" & _
            customers.Item(1).CustomerName
        .Range("E9").Value2 = CStr(filtered.Count) & "|" & CStr(filtered.Last)
        .Range("E10").Value2 = JoinList(projected, ",")
        .Range("E11").Value2 = JoinList(sequence, ",")
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
    Dim ageSelector As ROneCOne
    Dim allExperienced As Boolean
    Dim anyLondon As Boolean
    Dim cityPredicate As ROneCOne
    Dim customerPrototype As DemoCustomer
    Dim customers As ROneCOne
    Dim experienced As ROneCOne
    Dim firstCustomer As DemoCustomer
    Dim lastCustomer As DemoCustomer
    Dim minimumAgePredicate As ROneCOne
    Dim nameParameter As ROneCOne
    Dim names As ROneCOne
    Dim nameSelector As ROneCOne
    Dim orderedCustomers As ROneCOne
    Dim query As DemoCustomerQuery

    Set customerPrototype = New DemoCustomer
    Set customers = ROneCOne.ListOf(customerPrototype)
    Set customerPrototype = Nothing

    customers.Add CreateCustomer("Ada", CLng(36), "London")
    customers.Add CreateCustomer("Grace", CLng(40), "Arlington")
    customers.Add CreateCustomer("Katherine", CLng(49), "Cleveland")

    Set query = New DemoCustomerQuery
    query.MinimumAge = 40
    query.RequiredCity = "London"
    Set minimumAgePredicate = ROneCOne.FromMethod( _
        query, "MeetsMinimumAge", 1)
    Set cityPredicate = ROneCOne.FromMethod(query, "IsInRequiredCity", 1)
    Set nameSelector = ROneCOne.FromMethod(query, "SelectName", 1)
    Set ageSelector = ROneCOne.FromMethod(query, "SelectAge", 1)

    ' Build the query first, then mutate its source to prove it remains deferred.
    Set experienced = customers.Where(minimumAgePredicate)
    customers.Add CreateCustomer("Margaret", CLng(45), "New York")
    Set lastCustomer = experienced.Last

    ' Select the customer names and order the resulting List<String>.
    Set nameParameter = ROneCOne.Parameter(vbString)
    Set names = experienced _
        .SelectItems(nameSelector, vbString) _
        .OrderBy(ROneCOne.Lambda(nameParameter, nameParameter)) _
        .ToList

    ' Ordering objects preserves T, so First returns a DemoCustomer instance.
    Set orderedCustomers = customers.OrderByDescending(ageSelector).ToList
    Set firstCustomer = orderedCustomers.First

    anyLondon = customers.AnyItem(cityPredicate)
    allExperienced = customers.All(minimumAgePredicate)

    With ThisWorkbook.Worksheets(USER_CLASS_SHEET)
        .Range("E6").Value2 = customers.GenericTypeName & ":" & customers.Count
        .Range("E7").Value2 = CStr(experienced.Count) & "|" & _
            lastCustomer.CustomerName
        .Range("E8").Value2 = JoinList(names, "|")
        .Range("E9").Value2 = firstCustomer.CustomerName & "|" & _
            CStr(firstCustomer.Age)
        .Range("E10").Value2 = CStr(anyLondon) & "|" & CStr(allExperienced)
        .Range("E11").Value2 = Round(CDbl( _
            experienced.SelectItems(ageSelector, vbLong).Average), 1)
    End With
End Sub

' -----------------------------------------------------------------------------
' Benchmark and reporting
' -----------------------------------------------------------------------------

Private Sub RunCollectionBenchmark()
    Dim filtered As ROneCOne
    Dim numbers As ROneCOne
    Dim started As Double
    Dim x As ROneCOne

    Set x = ROneCOne.Parameter(vbLong)
    started = Timer
    Set numbers = ROneCOne.Range(CLng(1), BENCHMARK_ELEMENT_COUNT)
    Set filtered = numbers _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .ToList

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = filtered.Count
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

Private Function JoinList(ByVal values As ROneCOne, ByVal separator As String) As String
    Dim index As Long
    Dim result As String

    For index = 0 To values.Count - 1
        If index > 0 Then result = result & separator
        result = result & CStr(values.Item(index))
    Next index
    JoinList = result
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
