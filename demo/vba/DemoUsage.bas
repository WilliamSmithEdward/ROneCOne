Attribute VB_Name = "DemoUsage"
Option Explicit

Public Sub RunROneCOneDemo()
    Dim x As ROneCOne
    Dim y As ROneCOne
    Dim square As ROneCOne
    Dim addValues As ROneCOne
    Dim between As ROneCOne
    Dim safeFalse As ROneCOne
    Dim maximum As ROneCOne
    Dim doubleValue As ROneCOne
    Dim pipeline As ROneCOne
    Dim worksheetFunctions As Object
    Dim started As Double
    Dim index As Long
    Dim lastResult As Variant
    Dim capturedNumber As Long
    Dim capturedDescription As String

    On Error GoTo DemoFailure

    Set x = ROneCOne.Parameter(vbLong)
    Set y = ROneCOne.Parameter(vbLong)
    Set square = ROneCOne.Lambda(x.Multiply(x), x)
    Set addValues = ROneCOne.Lambda(x.Add(y), x, y)
    Set between = ROneCOne.Lambda( _
        x.GreaterThan(CLng(10)).AndAlso(x.LessThan(CLng(20))), _
        x)
    Set safeFalse = ROneCOne.Lambda( _
        ROneCOne.Value(False).AndAlso(ROneCOne.Value(1).Divide(0)))
    Set worksheetFunctions = Application.WorksheetFunction
    Set maximum = ROneCOne.FromMethod(worksheetFunctions, "Max", 2)
    Set doubleValue = ROneCOne.Lambda(x.Add(x), x)
    Set pipeline = square.PipeTo(doubleValue)

    With ThisWorkbook.Worksheets("Examples")
        .Range("E6").Value2 = square(CLng(9))
        .Range("E7").Value2 = addValues(CLng(6), CLng(7))
        .Range("E8").Value2 = between(CLng(15))
        .Range("E9").Value2 = safeFalse.Run()
        .Range("E10").Value2 = maximum(CLng(4), CLng(7))
        .Range("E11").Value2 = pipeline(CLng(3))
    End With

    started = Timer
    For index = 1 To 10000
        lastResult = square(CLng(index))
    Next index

    With ThisWorkbook.Worksheets("Benchmarks")
        .Range("B6").Value2 = 10000
        .Range("C6").Value2 = Timer - started
        .Range("D6").Value2 = lastResult
    End With

    With ThisWorkbook.Worksheets("Start Here")
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "PASS"
        .Range("B14").ClearContents
    End With
    Application.Calculate
    Exit Sub

DemoFailure:
    capturedNumber = Err.Number
    capturedDescription = Err.Description
    With ThisWorkbook.Worksheets("Start Here")
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "ERROR"
        .Range("B14").Value2 = CStr(capturedNumber) & ": " & capturedDescription
    End With
End Sub
