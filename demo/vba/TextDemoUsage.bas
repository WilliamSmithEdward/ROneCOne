Attribute VB_Name = "TextDemoUsage"
Option Explicit

' ============================================================================
' ROneCOne tutorial: regular expressions, hashing, and encoding
' ----------------------------------------------------------------------------
' This demo never touches the network. It matches text with real regular
' expressions, computes the same digests every other platform produces, and
' moves bytes between base64 and hexadecimal, all in-process.
'
' The surface mirrors what C# programmers know: ROneCOne.Regex has the
' System.Text.RegularExpressions verbs (IsMatch, Match, Matches, Replace,
' Split); ROneCOne.Hash computes SHA and HMAC digests through Windows CNG;
' and ROneCOne.Convert encodes byte arrays as base64 or hex. Text hashes as
' its UTF-8 bytes, so a digest here equals the one from Python or sha256sum.
'
' To run it: press Alt+F8, choose RunROneCOneTextDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneTextDemo()
    Dim errorDescription As String
    Dim errorNumber As Long

    On Error GoTo DemoFailure
    WriteTextExamples
    RunTextBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errorNumber, errorDescription
End Sub

Private Sub WriteTextExamples()
    Dim email As ROneCOne

    ' A compiled pattern is reusable. This one captures the local part, the
    ' domain, and the top-level suffix of an email address as three groups.
    Set email = ROneCOne.Regex("(\w+)@(\w+)\.(\w+)")

    ' Each line reads one result and writes it to the Examples sheet, so
    ' every feature shows its answer next to what the sheet expects. The
    ' digests are the published FIPS 180 and RFC 4231 test vectors.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = email.IsMatch("write ada@x.com today")
        .Range("E7").Value2 = _
            email.Match("write ada@x.com today").Groups.Item(1)
        .Range("E8").Value2 = email.Matches("ada@x.com and bo@y.org").Count
        .Range("E9").Value2 = email.Replace("ada@x.com", "$2:$1")
        .Range("E10").Value2 = ROneCOne.Regex("\s*,\s*").Split("a, b ,c").Count
        .Range("E11").Value2 = _
            ROneCOne.Convert.ToHexString(ROneCOne.Hash.Sha256("abc"))
        .Range("E12").Value2 = ROneCOne.Convert.ToHexString( _
            ROneCOne.Hash.HmacSha256("Jefe", _
            "what do ya want for nothing?"))
        .Range("E13").Value2 = ROneCOne.Convert.ToBase64String( _
            ROneCOne.Convert.FromHexString("4D616E"))
        .Range("E14").Value2 = ROneCOne.Convert.ToHexString( _
            ROneCOne.Convert.FromHexString("4d616e"))
    End With
End Sub

Private Sub RunTextBenchmark()
    Dim digests As ROneCOne
    Dim elapsed As Double
    Dim index As Long
    Dim started As Double

    ' Hash a thousand distinct strings and count the distinct digests. CNG
    ' runs at native speed, so the whole batch stays well under a second.
    Set digests = ROneCOne.HashSetOf(vbString)
    started = Timer
    For index = 1 To BENCHMARK_ROWS
        digests.Add ROneCOne.Convert.ToHexString( _
            ROneCOne.Hash.Sha256("row-" & CStr(index)))
    Next index
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = digests.Count
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

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function
