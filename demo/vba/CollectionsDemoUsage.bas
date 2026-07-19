Attribute VB_Name = "CollectionsDemoUsage"
Option Explicit

Public Sub RunROneCOneCollectionsDemo()
    Dim actualError As Long
    Dim ada As DemoCustomer
    Dim collectionSeconds As Double
    Dim collectionStarted As Double
    Dim customers As ROneCOne
    Dim customerPrototype As DemoCustomer
    Dim enumerationValues As ROneCOne
    Dim filtered As ROneCOne
    Dim grace As DemoCustomer
    Dim numbers As ROneCOne
    Dim projected As ROneCOne
    Dim sequence As ROneCOne
    Dim total As Long
    Dim value As Variant
    Dim x As ROneCOne

    On Error GoTo DemoFailure

    Set numbers = ROneCOne.ListOf(vbLong)
    numbers.Add CLng(5)
    numbers.Add CLng(10)
    numbers.Add CLng(15)

    On Error Resume Next
    numbers.Add "not a Long"
    actualError = Err.Number
    Err.Clear
    On Error GoTo DemoFailure

    Set customerPrototype = New DemoCustomer
    Set customers = ROneCOne.ListOf(customerPrototype)
    Set ada = New DemoCustomer
    ada.CustomerName = "Ada"
    Set grace = New DemoCustomer
    grace.CustomerName = "Grace"
    customers.Add ada
    customers.Add grace

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

    With ThisWorkbook.Worksheets("Examples")
        .Range("E6").Value2 = numbers.GenericTypeName
        .Range("E7").Value2 = (actualError = ROneCOne.TypeMismatchError)
        .Range("E8").Value2 = customers.GenericTypeName & ":" & _
            customers.Item(1).CustomerName
        .Range("E9").Value2 = CStr(filtered.Count) & "|" & _
            CStr(filtered.Last)
        .Range("E10").Value2 = CStr(projected.Item(0)) & "," & _
            CStr(projected.Item(1))
        .Range("E11").Value2 = CStr(sequence.Item(0)) & "," & _
            CStr(sequence.Item(1)) & "," & CStr(sequence.Item(2))
        .Range("E12").Value2 = CStr(numbers.Sum) & "|" & _
            CStr(numbers.Average) & "|" & CStr(numbers.Min) & "|" & _
            CStr(numbers.Max)
        .Range("E13").Value2 = total
    End With

    collectionStarted = Timer
    Set numbers = ROneCOne.Range(CLng(1), CLng(10000))
    Set filtered = numbers _
        .Where(ROneCOne.Lambda(x.Modulo(CLng(2)).EqualTo(CLng(0)), x)) _
        .ToList
    collectionSeconds = ElapsedSeconds(collectionStarted)
    With ThisWorkbook.Worksheets("Benchmarks")
        .Range("B6").Value2 = 10000
        .Range("C6").Value2 = collectionSeconds
        .Range("D6").Value2 = filtered.Count
    End With

    With ThisWorkbook.Worksheets("Start Here")
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "PASS"
        .Range("B14").ClearContents
    End With
    Application.Calculate
    Exit Sub

DemoFailure:
    With ThisWorkbook.Worksheets("Start Here")
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "ERROR"
        .Range("B14").Value2 = CStr(Err.Number) & ": " & Err.Description
    End With
End Sub

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
