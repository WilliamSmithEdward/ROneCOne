Attribute VB_Name = "CollectionsDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: collections and LINQ-style queries
' ----------------------------------------------------------------------------
' This is the broad tour. It shows how to hold data in a typed list and then
' ask questions of it: filter to the rows you want, reshape them, sort them,
' and summarize them, without writing loops, counters, or temporary arrays.
'
' The queries read almost like sentences. "Where Age at least 40" keeps the
' rows whose Age is 40 or more; "Map CustomerName" pulls out just the names;
' "OrderBy City" sorts them. A query is also lazy: it does no work until you
' ask for the result (for example with ToList), so it always sees the latest
' data, even rows added after the query was written.
'
' The demo works with plain numbers first, then with ordinary Customer objects
' of your own class, to show the same queries apply to both. Read the public
' macro, then each small procedure top to bottom.
'
' To run it: press Alt+F8, choose RunROneCOneCollectionsDemo, and click Run.
' Results land on the "Examples" and "User Class LINQ" worksheets.
' ============================================================================

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
' Checked number lists and simple queries
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

    ' A typed list holds one kind of value. This one holds Long whole numbers.
    ' Because it is typed, adding the wrong kind of value is refused rather than
    ' silently stored; RaisesExpectedTypeMismatch confirms that guard fires.
    Set numbers = ROneCOne.ListOf( _
        vbLong, 5, 10, 15)
    strictTypeRejected = RaisesExpectedTypeMismatch(numbers)

    ' ListFrom builds a list from objects you already have and figures out the
    ' element type from the first one, so this becomes a list of DemoCustomer.
    Set ada = CreateCustomer("Ada", 36, "London")
    Set grace = CreateCustomer("Grace", 40, "Arlington")
    Set customers = ROneCOne.ListFrom(ada, grace)

    ' Queries are lazy. "filtered" is defined here as "numbers greater than 10",
    ' but it does not actually look at the numbers until ToList asks for them on
    ' the last line. By then 30 has been added, so the result includes it. The
    ' query describes what you want; ToList is the moment it runs. Element is a
    ' stand-in for "each number" while the condition is being described.
    Set numbers = ROneCOne.ListOf(vbLong, 5, 20)
    Set x = numbers.Element
    Set filtered = numbers.Where(x.GreaterThan(10))
    numbers.Add 30
    Set filtered = filtered.ToList

    ' Query steps chain left to right, each feeding the next. Read this as: take
    ' the numbers 1 through 6, keep the even ones, multiply each by 10, sort
    ' high to low, and keep the first two. Range(1, 6) means "6 numbers from 1".
    Set projected = ROneCOne.Range(1, 6) _
        .Where(x.Modulo(2).EqualTo(0)) _
        .Map(x.Multiply(10), vbLong) _
        .OrderDescending _
        .Take(2) _
        .ToList

    ' The same chaining shapes a sequence: remove duplicates, add a value at the
    ' front and the back, reverse the order, then drop the first item.
    Set sequence = ROneCOne.ListOf(vbLong, 2, 2, 3)
    Set sequence = sequence.Distinct _
        .Prepend(1) _
        .Append(4) _
        .Reverse _
        .Skip(1) _
        .ToList

    ' ForEach runs one action against every element. Here it adds each number
    ' 1 through 4 into a running total (1 + 2 + 3 + 4 = 10) via a small handler.
    Set numbers = ROneCOne.Range(1, 5)
    Set enumerationValues = ROneCOne.Range(1, 4)
    mForEachTotal = 0
    enumerationValues.ForEach ROneCOne.Action( _
        "CollectionsDemoUsage.DemoAccumulateLong").Takes(vbLong)
    total = mForEachTotal

    With ThisWorkbook.Worksheets(BASIC_EXAMPLES_SHEET)
        .Range("E6").Value2 = numbers.GenericTypeName
        .Range("E7").Value2 = strictTypeRejected
        .Range("E8").Value2 = customers.GenericTypeName & _
            "; second customer: " & customers.Item(1).CustomerName
        .Range("E9").Value2 = CStr(filtered.Count) & _
            " matches; last: " & CStr(filtered.Last)
        .Range("E10").Value2 = "Top results: " & projected.JoinText(", ")
        .Range("E11").Value2 = "Sequence: " & sequence.JoinText(", ")
        .Range("E12").Value2 = "Sum " & CStr(numbers.Sum) & _
            "; average " & CStr(numbers.Average) & _
            "; min " & CStr(numbers.Min) & "; max " & CStr(numbers.Max)
        .Range("E13").Value2 = total
    End With
End Sub

' -----------------------------------------------------------------------------
' Query ordinary Customer objects
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

    ' The same queries now run over ordinary objects of your own class. Nothing
    ' about DemoCustomer knows about ROneCOne; you write a normal class and query
    ' it by property name. Each customer here also has a Manager and, later,
    ' a list of direct Reports, so the demo can show nested queries too.
    Set ada = CreateCustomer("Ada", 36, "London")
    Set grace = CreateCustomer("Grace", 40, "Arlington")
    Set grace.Manager = ada
    Set katherine = CreateCustomer("Katherine", 49, "Cleveland")
    Set katherine.Manager = grace
    Set customers = ROneCOne.ListFrom(ada, grace, katherine)

    ' Build the query now; it will include Margaret when results are requested.
    Set experienced = customers.Where("Age").AtLeast(40)
    Set margaret = CreateCustomer("Margaret", 45, "Arlington")
    Set margaret.Manager = grace
    customers.Add margaret
    Set lastCustomer = experienced.Last

    ' Use the property name directly instead of writing a selector procedure.
    Set names = experienced _
        .Map("CustomerName", vbString) _
        .Order _
        .ToList

    ' Sort by city first, then by age within each city.
    Set orderedCustomers = customers _
        .OrderBy("City") _
        .ThenByDescending("Age") _
        .ToList
    Set firstCustomer = orderedCustomers.First

    ' A Condition is a reusable yes-or-no rule you can name once and pass around.
    ' Exists is true if any customer matches; All is true only if every one does.
    anyLondon = customers.Exists( _
        customers.Condition("City").EqualTo("London"))
    allExperienced = customers.All( _
        customers.Condition("Age").AtLeast(40))

    ' An existing VBA function can also act as the yes-or-no rule.
    Set procedureFiltered = customers.WhereMethod( _
        "CollectionsDemoUsage.IsExperiencedCustomer").ToList

    Set stringMatches = customers _
        .Where("CustomerName").StartsWith("G") _
        .ToList
    Set distinctCities = customers.DistinctBy("City").ToList

    ' The "?." in "Manager?.Age" is a safety valve: if a customer has no Manager,
    ' the path stops instead of erroring, and that customer just does not match.
    Set managed = customers.Where("Manager?.Age").AtLeast(40).ToList

    ' When the built-in rules are not quite what you need, you can supply your
    ' own. An equality comparer decides when two values count as "the same"; a
    ' comparer decides their sort order. These two treat text case-insensitively,
    ' so "Ada" and "ADA" are one value here. SingleItem asks for the one and only
    ' customer that matches, and complains if there is not exactly one.
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

    ' Give each customer a list of direct reports, so a customer now contains a
    ' collection of other customers. That lets you ask questions about the inner
    ' list from the outer one (below): which customers have any report aged 40+,
    ' all reports past a threshold, or none matching, and it shows a filter can
    ' be driven by an expression comparing one field to a list of allowed values.
    Set ada.Reports = ROneCOne.ListFrom(grace, katherine)
    Set grace.Reports = ROneCOne.ListFrom(ada)
    Set katherine.Reports = ROneCOne.ListLike(ada)
    Set margaret.Reports = ROneCOne.ListFrom(grace)
    Set reportPredicate = ada.Reports.Condition("Age").AtLeast(40)
    Set membershipMatches = customers.Where(allowedCities.Contains(customers!City))

    With ThisWorkbook.Worksheets(USER_CLASS_SHEET)
        .Range("E6").Value2 = customers.GenericTypeName & " with " & _
            CStr(customers.Count) & " customers"
        .Range("E7").Value2 = CStr(experienced.Count) & _
            " customers; newest match: " & lastCustomer.CustomerName
        .Range("E8").Value2 = names.JoinText(", ")
        .Range("E9").Value2 = "First: " & firstCustomer.CustomerName & _
            ", age " & CStr(firstCustomer.Age)
        .Range("E10").Value2 = "London exists: " & CStr(anyLondon) & _
            "; all age 40+: " & CStr(allExperienced)
        .Range("E11").Value2 = Round(CDbl( _
            experienced.Average("Age")), 1)
        .Range("E12").Value2 = CStr(procedureFiltered.Count) & _
            " matches; " & _
            customers.Predicate( _
                "CollectionsDemoUsage.IsExperiencedCustomer").Signature
        .Range("E13").Value2 = "Starts with G: " & _
            CStr(stringMatches.Count) & "; contains ther: " & _
            CStr(customers.Where("CustomerName").Contains("ther").Count)
        .Range("E14").Value2 = distinctCities.Count
        With customers
            Set names = .Where(!Age.AtLeast(40)) _
                .Map("CustomerName", vbString) _
                .Order _
                .ToList
        End With
        .Range("E15").Value2 = names.JoinText(", ")
        .Range("E16").Value2 = "No manager: " & _
            CStr(customers.Where("Manager").IsNothing.Count) & _
            "; manager age 40+: " & CStr(managed.Count)
        .Range("E17").Value2 = "Allowed city matches: " & _
            CStr(customers.Where("City").IsIn(allowedCities).Count) & _
            "; expression matches: " & CStr(membershipMatches.Count)
        .Range("E18").Value2 = "City matches: " & _
            CStr(customers.Where("City") _
                .EqualToIgnoreCase("LONDON").Count) & _
            "; name contains: " & CStr(customers.Where("CustomerName") _
                .ContainsIgnoreCase("THER").Count)
        .Range("E19").Value2 = "Count: " & CStr(customers.Count( _
            customers.Condition("Age").AtLeast(40))) & _
            "; single: " & singleCustomer.CustomerName & _
            "; none age 100: " & _
            CStr(customers.None(customers.Match("Age", 100)))
        .Range("E20").Value2 = "Any: " & CStr(customers.WhereAny( _
            "Reports", reportPredicate).Count) & "; all: " & _
            CStr(customers.WhereAll("Reports", ada.Reports.Condition("Age") _
                .AtLeast(36)).Count) & "; none: " & _
            CStr(customers.WhereNone("Reports", reportPredicate).Count)
        .Range("E21").Value2 = "Distinct: " & CStr(strings.Distinct( _
            equalityComparer).Count) & "; contains Ada: " & _
            CStr(strings.Contains("ada", equalityComparer)) & _
            "; first: " & CStr(strings.Order(comparer).First)
        .Range("E22").Value2 = "Both: " & CStr(customers.Where( _
            customers.Condition("Age").AtLeast(40).Both( _
                customers.Match("City", "Arlington"))).Count) & _
            "; either: " & CStr(customers.Where(customers.Match( _
                "City", "London").Either(customers.Match( _
                    "City", "Cleveland"))).Count)
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

    ' The same operations, now at ten thousand elements, timed so you can see
    ' they stay fast enough for everyday workbook use: filter, then (below)
    ' object filter, multi-key sort, and dictionary build plus indexed lookup.
    started = Timer
    Set numbers = ROneCOne.Range(1, BENCHMARK_ELEMENT_COUNT)
    Set x = numbers.Element
    Set filtered = numbers _
        .Where(x.Modulo(2).EqualTo(0)) _
        .ToList

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C6").Value2 = ElapsedSeconds(started)
        .Range("D6").Value2 = filtered.Count
    End With

    Set customer = CreateCustomer("Benchmark", 50, "Local")
    Set customers = ROneCOne.Repeat(customer, BENCHMARK_ELEMENT_COUNT)
    started = Timer
    Set filtered = customers.Where("Age").AtLeast(40).ToList
    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B7").Value2 = BENCHMARK_ELEMENT_COUNT
        .Range("C7").Value2 = ElapsedSeconds(started)
        .Range("D7").Value2 = filtered.Count
    End With

    started = Timer
    Set ordered = numbers _
        .OrderBy(x.Modulo(100)) _
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
