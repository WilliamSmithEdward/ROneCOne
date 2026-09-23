Attribute VB_Name = "FilesDemoUsage"
Option Explicit

' ROneCOne 1.10.1, released 2026-09-23
'
' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' ============================================================================
' ROneCOne tutorial: files, folders, and CSV in the spirit of System.IO
' ----------------------------------------------------------------------------
' This demo never touches the network. Everything it writes lives under one
' folder next to this workbook, and the folder is removed at the end, so the
' demo leaves nothing behind. The point is that VBA can finally speak the
' file formats other tools expect: UTF-8 text without byte-order marks,
' folders created and enumerated like System.IO, and CSV that round-trips a
' typed table without a single Split or hand-rolled quote handler.
'
' The surface mirrors what C# programmers know: File.ReadAllText and
' WriteAllText, Directory.CreateDirectory and GetFiles, Path.Combine, and a
' Csv serializer that pairs with table.ToCsv the way ToJson pairs with the
' JSON layer.
'
' To run it: press Alt+F8, choose RunROneCOneFilesDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneFilesDemo()
    Dim demoRoot As String
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    demoRoot = ThisWorkbook.Path & "\ROneCOne_Files_Demo_Data"
    If ROneCOne.Directory.Exists(demoRoot) Then
        ROneCOne.Directory.Delete demoRoot, True
    End If
    ROneCOne.Directory.CreateDirectory demoRoot & "\inner"
    WriteFileExamples demoRoot
    RunFilesBenchmark demoRoot
    ROneCOne.Directory.Delete demoRoot, True
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Len(demoRoot) > 0 Then
        If ROneCOne.Directory.Exists(demoRoot) Then
            ROneCOne.Directory.Delete demoRoot, True
        End If
    End If
    On Error GoTo 0
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteFileExamples(ByVal demoRoot As String)
    Dim readBack As ROneCOne
    Dim orders As ROneCOne
    Dim roundTripped As ROneCOne

    ' Step 1: plain text, correctly encoded. WriteAllText produces UTF-8
    ' without a byte-order mark, exactly what git, Python, and web services
    ' expect; ReadAllText hands the same characters back.
    ROneCOne.File.WriteAllText demoRoot & "\hello.txt", "hello files"

    ' Step 2: byte-order marks decide decoding on read. This file is UTF-16,
    ' but ReadAllText was not told that; the mark at the front of the file
    ' overrides the default and the text still comes back intact.
    ROneCOne.File.WriteAllText demoRoot & "\hello16.txt", _
        "hello files", "utf-16"

    ' Step 3: lines are a collection. WriteAllLines terminates every line,
    ' and ReadAllLines returns an ordinary typed list you can query.
    ROneCOne.File.WriteAllLines demoRoot & "\lines.txt", _
        Array("alpha", "beta", "gamma")
    Set readBack = ROneCOne.File.ReadAllLines(demoRoot & "\lines.txt")

    ' Step 4: a typed table becomes a CSV file in two calls, and the file
    ' becomes a typed table again in two more. The Note column shows the
    ' round trip preserving a database null (an empty field) as a null.
    Set orders = ROneCOne.DataTable("Orders")
    orders.Column "Id", vbLong
    orders.Column "Customer", vbString
    orders.Column "Total", vbDouble
    orders.Column "Note", vbString
    orders.LoadRow Array(1, "Ada", 12.5, ROneCOne.DBNull)
    orders.LoadRow Array(2, "Bo", 20#, "rush")
    orders.LoadRow Array(3, "Cy", 7.25, "gift")
    ROneCOne.File.WriteAllText demoRoot & "\orders.csv", orders.ToCsv
    Set roundTripped = ROneCOne.Csv.DeserializeTable( _
        ROneCOne.File.ReadAllText(demoRoot & "\orders.csv"), "Orders")

    ' Step 5: folders behave like System.IO.Directory. One extra file sits
    ' in a subfolder so the recursive enumeration below has something to
    ' find beyond the top level.
    ROneCOne.File.WriteAllText demoRoot & "\inner\note.txt", "inner note"

    ' Step 6: a logger writes level-coded, timestamped lines. The Debug line
    ' is below the default Information level, so only two lines land.
    Dim logLines As ROneCOne
    Dim runLog As ROneCOne
    Set runLog = ROneCOne.Logger(demoRoot & "\run.log")
    runLog.LogInformation "processed {0} rows", 3
    runLog.LogDebug "this line is filtered out"
    runLog.LogWarning "keep an eye on row {0}", 2
    Set logLines = ROneCOne.File.ReadAllLines(demoRoot & "\run.log")

    ' Step 7: a watcher awaits the next change under a folder. The baseline
    ' is taken now, so the file dropped just below is seen as Created.
    Dim fileChange As ROneCOne
    Dim watchTask As ROneCOne
    Dim watcher As ROneCOne
    Set watcher = ROneCOne.FileWatcher(demoRoot, "*.dat")
    Set watchTask = watcher.WaitForChangeAsync
    ROneCOne.File.WriteAllText demoRoot & "\signal.dat", "ready"
    Set fileChange = watchTask.Await

    ' Each line reads one result and writes it to the Examples sheet, so
    ' every feature above shows its answer next to what the sheet expects.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = ROneCOne.File.ReadAllText(demoRoot & "\hello.txt")
        .Range("E7").Value2 = (ROneCOne.File.ReadAllText( _
            demoRoot & "\hello16.txt") = "hello files")
        .Range("E8").Value2 = readBack.Count
        .Range("E9").Value2 = ROneCOne.Path.Combine("C:\data", "in", "file.txt")
        .Range("E10").Value2 = ROneCOne.Path.GetFileNameWithoutExtension( _
            "C:\data\in\file.txt")
        .Range("E11").Value2 = ROneCOne.Directory.GetFiles( _
            demoRoot, "*.txt", True).Count
        .Range("E12").Value2 = ROneCOne.File.Exists(demoRoot & "\orders.csv")
        .Range("E13").Value2 = roundTripped.Rows.Count
        .Range("E14").Value2 = roundTripped.Rows.Item(0).Item("Total")
        .Range("E15").Value2 = IsNull(roundTripped.Rows.Item(0).Item("Note"))
        .Range("E16").Value2 = logLines.Count
        .Range("E17").Value2 = fileChange.ChangeType & " " & fileChange.Name
    End With
End Sub

Private Sub RunFilesBenchmark(ByVal demoRoot As String)
    Dim csvPath As String
    Dim elapsed As Double
    Dim idx As Long
    Dim roundTripped As ROneCOne
    Dim started As Double
    Dim benchTable As ROneCOne

    ' One thousand typed rows travel out to a CSV file on disk and back into
    ' a fresh typed table: serialize, write, read, and parse, all timed as
    ' one round trip.
    Set benchTable = ROneCOne.DataTable("Benchmark")
    benchTable.Column "Id", vbLong
    benchTable.Column "Customer", vbString
    benchTable.Column "Total", vbDouble
    For idx = 1 To BENCHMARK_ROWS
        benchTable.LoadRow Array(idx, "C" & CStr(idx), idx * 1.5)
    Next idx
    csvPath = demoRoot & "\benchmark.csv"

    started = Timer
    ROneCOne.File.WriteAllText csvPath, benchTable.ToCsv
    Set roundTripped = ROneCOne.Csv.DeserializeTable( _
        ROneCOne.File.ReadAllText(csvPath), "Benchmark")
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = roundTripped.Rows.Count
    End With
End Sub

Private Sub MarkDemoPassed()
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "PASS"
        .Range("B14").ClearContents
    End With
End Sub

Private Sub MarkDemoFailed(ByVal errNumber As Long, ByVal errDescription As String)
    With ThisWorkbook.Worksheets(START_SHEET)
        .Range("B12").Value2 = Now
        .Range("B13").Value2 = "ERROR"
        .Range("B14").Value2 = CStr(errNumber) & ": " & errDescription
    End With
End Sub

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
