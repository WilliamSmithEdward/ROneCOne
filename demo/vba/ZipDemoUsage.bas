Attribute VB_Name = "ZipDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: zip archives in the spirit of System.IO.Compression
' ----------------------------------------------------------------------------
' This demo never touches the network. Everything it writes lives under one
' folder next to this workbook, and the folder is removed at the end, so the
' demo leaves nothing behind. The point is that VBA can finally open and make
' zip archives with no reference, no add-in, and no Shell automation: a pure
' arithmetic engine parses the archive, inflates DEFLATE streams, and checks
' every entry's CRC-32.
'
' The surface mirrors what C# programmers know: ZipFile.OpenRead with Entries
' and GetEntry, entry.ReadAllText without extracting to disk,
' CreateFromDirectory and ExtractToDirectory. Interop is proven both ways
' against Windows PowerShell: the engine inflates what Compress-Archive
' deflates, and Expand-Archive reads what the engine writes.
'
' To run it: press Alt+F8, choose RunROneCOneZipDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneZipDemo()
    Dim demoRoot As String
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    demoRoot = ThisWorkbook.Path & "\ROneCOne_Zip_Demo_Data"
    If ROneCOne.Directory.Exists(demoRoot) Then
        ROneCOne.Directory.Delete demoRoot, True
    End If
    ROneCOne.Directory.CreateDirectory demoRoot
    WriteZipExamples demoRoot
    RunZipBenchmark demoRoot
    ROneCOne.Directory.Delete demoRoot, True
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Len(demoRoot) > 0 Then
        If ROneCOne.Directory.Exists(demoRoot) Then
            ROneCOne.Directory.Delete demoRoot, True
        End If
    End If
    On Error GoTo 0
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteZipExamples(ByVal demoRoot As String)
    Dim archive As ROneCOne
    Dim madePath As String
    Dim psMadePath As String
    Dim slipBytes As Variant
    Dim slipPath As String
    Dim slipTrace As String
    Dim sourceRoot As String
    Dim index As Long

    ' Step 1: build a small tree and zip it. CreateFromDirectory writes a
    ' stored archive that every unzipper reads, including PowerShell's own.
    sourceRoot = ROneCOne.Path.Combine(demoRoot, "source")
    ROneCOne.Directory.CreateDirectory ROneCOne.Path.Combine(sourceRoot, "sub")
    ROneCOne.File.WriteAllText _
        ROneCOne.Path.Combine(sourceRoot, "readme.txt"), "read me first"
    ROneCOne.File.WriteAllText ROneCOne.Path.Combine( _
        ROneCOne.Path.Combine(sourceRoot, "sub"), "data.csv"), _
        "id,total" & vbCrLf & "1,2.5"
    madePath = ROneCOne.Path.Combine(demoRoot, "made.zip")
    ROneCOne.ZipFile.CreateFromDirectory sourceRoot, madePath

    ' Step 2: open the archive and read an entry without extracting it. The
    ' entry knows its own name, size, and content.
    Set archive = ROneCOne.ZipFile.OpenRead(madePath)

    ' Step 3: interop. PowerShell zips the same tree with real DEFLATE
    ' compression; the engine opens it and inflates an entry. Then the
    ' engine's own archive is handed to Expand-Archive.
    Shell80 "Compress-Archive -LiteralPath '" & ROneCOne.Path.Combine( _
        sourceRoot, "readme.txt") & "','" & ROneCOne.Path.Combine( _
        sourceRoot, "sub") & "' -DestinationPath '" & _
        ROneCOne.Path.Combine(demoRoot, "ps_made.zip") & "'"
    psMadePath = ROneCOne.Path.Combine(demoRoot, "ps_made.zip")
    Shell80 "Expand-Archive -Path '" & madePath & "' -DestinationPath '" & _
        ROneCOne.Path.Combine(demoRoot, "ps_out") & "'"

    ' Step 4: the traversal guard. A hostile entry name is spliced into a
    ' real archive; extraction refuses it before writing anything.
    slipBytes = ROneCOne.File.ReadAllBytes(madePath)
    For index = 0 To UBound(slipBytes) - 9
        If BytesMatch(slipBytes, index, "readme.txt") Then
            PatchBytes slipBytes, index, "..\evil.x"
        End If
    Next index
    slipPath = ROneCOne.Path.Combine(demoRoot, "hostile.zip")
    ROneCOne.File.WriteAllBytes slipPath, slipBytes
    slipTrace = "guard did not fire"
    On Error Resume Next
    ROneCOne.ZipFile.ExtractToDirectory slipPath, _
        ROneCOne.Path.Combine(demoRoot, "guarded")
    If Err.Number = ROneCOne.ZipError Then slipTrace = "traversal refused"
    On Error GoTo 0

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = archive.Entries.Count
        .Range("E7").Value2 = archive.GetEntry("readme.txt").ReadAllText()
        .Range("E8").Value2 = archive.GetEntry("sub/data.csv").Length
        .Range("E9").Value2 = (archive.GetEntry("sub/data.csv") _
            .CompressedLength = archive.GetEntry("sub/data.csv").Length)
        .Range("E10").Value2 = (ROneCOne.ZipFile.OpenRead(psMadePath) _
            .Entries.Count >= 2)
        .Range("E11").Value2 = InStr(1, ROneCOne.ZipFile.OpenRead( _
            psMadePath).GetEntry("readme.txt").ReadAllText(), _
            "read me first") > 0
        .Range("E12").Value2 = ROneCOne.File.ReadAllText( _
            ROneCOne.Path.Combine(ROneCOne.Path.Combine( _
            demoRoot, "ps_out"), "readme.txt"))
        .Range("E13").Value2 = slipTrace
    End With
End Sub

Private Sub RunZipBenchmark(ByVal demoRoot As String)
    Dim archive As ROneCOne
    Dim builder As ROneCOne
    Dim elapsed As Double
    Dim index As Long
    Dim total As Double
    Dim zipPath As String

    ' PowerShell deflates a thousand-line file; the pure-VBA engine inflates
    ' it and sums the numbers back out, timing the read and inflate.
    Set builder = ROneCOne.StringBuilder()
    For index = 1 To BENCHMARK_ROWS
        builder.AppendLine CStr(index)
    Next index
    ROneCOne.File.WriteAllText _
        ROneCOne.Path.Combine(demoRoot, "nums.txt"), builder.ToString
    zipPath = ROneCOne.Path.Combine(demoRoot, "nums.zip")
    Shell80 "Compress-Archive -LiteralPath '" & ROneCOne.Path.Combine( _
        demoRoot, "nums.txt") & "' -DestinationPath '" & zipPath & "'"

    Dim started As Double
    started = Timer
    Set archive = ROneCOne.ZipFile.OpenRead(zipPath)
    Dim line As Variant
    For Each line In VBA.Split(archive.GetEntry("nums.txt").ReadAllText(), _
        vbCrLf)
        If Len(Trim$(CStr(line))) > 0 Then total = total + Val(CStr(line))
    Next line
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = total
    End With
End Sub

Private Sub Shell80(ByVal command As String)
    Dim result As ROneCOne

    Set result = ROneCOne.Process.RunAsync( _
        "powershell -NoProfile -Command """ & command & """").Await
    If result.ExitCode <> 0 Then
        Err.Raise vbObjectError + 5000, "ZipDemo", _
            "PowerShell step failed: " & result.StandardError
    End If
End Sub

Private Function BytesMatch( _
    ByRef bytes As Variant, _
    ByVal at As Long, _
    ByVal text As String _
) As Boolean
    Dim index As Long

    For index = 1 To Len(text)
        If bytes(at + index - 1) <> Asc(Mid$(text, index, 1)) Then
            Exit Function
        End If
    Next index
    BytesMatch = True
End Function

Private Sub PatchBytes( _
    ByRef bytes As Variant, _
    ByVal at As Long, _
    ByVal text As String _
)
    Dim index As Long

    For index = 1 To Len(text)
        bytes(at + index - 1) = Asc(Mid$(text, index, 1))
    Next index
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

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
