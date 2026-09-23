Attribute VB_Name = "HttpDemoUsage"
Option Explicit

' ROneCOne 1.10.0, released 2026-09-22
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
' ROneCOne tutorial: HTTP requests you can await
' ----------------------------------------------------------------------------
' This demo downloads live data from the internet without adding a single
' reference or add-in. You create an HttpClient, ask it for a URL, and get a
' task back; Await collects the response when it arrives. While the bytes are
' in flight, Windows does the downloading, so Excel stays responsive and your
' code stays ordinary VBA.
'
' The surface mirrors what C# programmers know from System.Net.Http: GetAsync
' for a full response, GetStringAsync for just the text, WhenAll to overlap
' several downloads, EnsureSuccessStatusCode to fail loudly, and a typed
' HttpRequestException you handle at the await site, the same place a C#
' try / await / catch would.
'
' The demo talks to https://pokeapi.co, a free public API about Pokemon.
' Nothing else is contacted, and the runtime itself never transmits anything.
' Running it needs an internet connection. The last two examples hand the
' downloaded JSON to ROneCOne.Json, turning it into a navigable tree and a
' typed DataTable without any extra library.
'
' To run it: press Alt+F8, choose RunROneCOneHttpDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_REQUESTS As Long = 3
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Private mTrace As String

Public Sub RunROneCOneHttpDemo()
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    WriteHttpExamples
    RunHttpBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteHttpExamples()
    Dim abilities As ROneCOne
    Dim berry As ROneCOne
    Dim canceledError As Long
    Dim client As ROneCOne
    Dim downloadPath As String
    Dim downloadWorked As Boolean
    Dim dittoJson As String
    Dim missError As Long
    Dim pikachuJson As String
    Dim responses As ROneCOne
    Dim response As ROneCOne
    Dim cancelSource As ROneCOne
    Dim pending As ROneCOne
    Dim tree As ROneCOne

    ' Step 1: create the client once and give it a base address, exactly like
    ' HttpClient in C#. Every request below can then use a short relative URL.
    ' The timeout says no single request may take longer than twenty seconds.
    Set client = ROneCOne.HttpClient()
    client.BaseAddress = "https://pokeapi.co/api/v2/"
    client.DefaultRequestHeader "Accept", "application/json"
    client.Timeout = 20000

    ' Download one resource. GetAsync starts the request and hands back a
    ' task immediately; Await collects the full response when it arrives.
    ' While the download is in flight, Windows moves the bytes and Excel
    ' keeps breathing.
    Set response = client.GetAsync("pokemon/pikachu").Await

    ' Sometimes you only want the text. GetStringAsync skips the response
    ' object and resolves straight to the body, and it refuses to hand back
    ' text from a failed request.
    dittoJson = client.GetStringAsync("pokemon/ditto").Await

    ' Overlap three downloads. Each GetAsync starts its transfer right away,
    ' so all three are in flight together; WhenAll finishes when the last one
    ' lands. The Benchmarks sheet times the same downloads both ways.
    Set responses = ROneCOne.Task.WhenAll( _
        client.GetAsync("pokemon/bulbasaur"), _
        client.GetAsync("pokemon/charmander"), _
        client.GetAsync("pokemon/squirtle")).Await

    ' Failure is part of the surface. A missing resource is still a response
    ' with a status code you can inspect, and GetStringAsync turns a bad
    ' status into a typed, catchable error. Exactly like try / await / catch
    ' in C#, you handle that error at the await site: trap it and compare its
    ' number with ROneCOne.HttpRequestError. (The task also keeps the full
    ' story in IsFaulted and Exception, mirroring Task.Exception in .NET.)
    Set pending = client.GetStringAsync("pokemon/missingno")
    mTrace = "unexpected success"
    On Error Resume Next
    pending.Await
    missError = Err.Number
    On Error GoTo 0
    If missError = ROneCOne.HttpRequestError And pending.IsFaulted Then
        mTrace = "skipped a missing resource"
    End If

    ' Cancellation uses the same tokens as every other task. This token is
    ' canceled before the request is awaited, so the transfer is abandoned
    ' and the task reports IsCanceled instead of a result.
    Set cancelSource = ROneCOne.CancellationTokenSource
    cancelSource.Cancel
    Set pending = client.GetAsync("pokemon/eevee", cancelSource.Token)
    On Error Resume Next
    pending.Await
    canceledError = Err.Number
    On Error GoTo 0

    ' A relative URL rides on the client's BaseAddress, so switching between
    ' test and production servers is one assignment, not a find-and-replace.
    Set berry = client.GetAsync("berry/1").Await

    ' Downloaded JSON becomes usable data in two calls. Deserialize turns the
    ' text into a navigable tree you can read by name, and DeserializeTable
    ' lands an array of objects straight in a typed DataTable; from there,
    ' ToRange could put the whole thing on a worksheet.
    pikachuJson = client.GetStringAsync("pokemon/pikachu").Await
    Set tree = ROneCOne.Json.Deserialize(pikachuJson)
    Set abilities = ROneCOne.Json.DeserializeTable( _
        pikachuJson, "Abilities", "$.abilities")

    ' A response body can go straight to disk. DownloadFileAsync composes the
    ' byte-array download with the file layer, so the file is on disk once the
    ' task completes; the demo saves next to this workbook and deletes after.
    downloadPath = ThisWorkbook.Path & "\ROneCOne_Http_Download.json"
    client.DownloadFileAsync("pokemon/pikachu", downloadPath).Await
    downloadWorked = ROneCOne.File.Exists(downloadPath) And InStr(1, _
        ROneCOne.File.ReadAllText(downloadPath), """pikachu""") > 0
    ROneCOne.File.Delete downloadPath

    ' Every example writes its answer next to what the sheet expects.
    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = CStr(response.StatusCode) & " " & _
            response.ReasonPhrase
        .Range("E7").Value2 = (InStr(1, dittoJson, """ditto""") > 0)
        .Range("E8").Value2 = _
            (response.EnsureSuccessStatusCode Is response)
        .Range("E9").Value2 = response.Header("Content-Type")
        .Range("E10").Value2 = DescribeOverlappedReplies(responses)
        .Range("E11").Value2 = client.GetAsync( _
            "pokemon/missingno").Await.StatusCode
        .Range("E12").Value2 = mTrace
        .Range("E13").Value2 = (pending.IsCanceled And canceledError <> 0)
        .Range("E14").Value2 = (berry.StatusCode = 200)
        .Range("E15").Value2 = CStr(tree.Item("name"))
        .Range("E16").Value2 = (abilities.Rows.Count > 0)
        .Range("E17").Value2 = downloadWorked
        ' Build a query value safely: EscapeDataString percent-encodes the
        ' spaces and ampersands that would otherwise break the URL. This is
        ' pure local text work, no request involved.
        .Range("E18").Value2 = _
            ROneCOne.Uri.EscapeDataString("name=Ada & Bo")
    End With
End Sub

Private Function DescribeOverlappedReplies( _
    ByVal responses As ROneCOne _
) As String
    If responses.Item(0).StatusCode = 200 And _
        responses.Item(1).StatusCode = 200 And _
        responses.Item(2).StatusCode = 200 Then
        DescribeOverlappedReplies = "bulbasaur, charmander, squirtle ready"
    Else
        DescribeOverlappedReplies = "a download failed"
    End If
End Function

Private Sub RunHttpBenchmark()
    Dim client As ROneCOne
    Dim ignored As ROneCOne
    Dim endpoints As Variant
    Dim overlappedElapsed As Double
    Dim resource As Variant
    Dim sequentialElapsed As Double
    Dim started As Double

    ' The same three downloads, timed two ways. One after another, each
    ' request waits for the previous one to finish. Overlapped, all three are
    ' in flight at once inside WinHTTP while VBA waits, so the total can come
    ' down toward the slowest single download. How much that saves depends on
    ' network latency, and fast responses can take the same time either way.
    Set client = ROneCOne.HttpClient()
    client.BaseAddress = "https://pokeapi.co/api/v2/"
    endpoints = Array("pokemon/bulbasaur", "pokemon/charmander", _
        "pokemon/squirtle")

    started = Timer
    Set ignored = ROneCOne.Task.WhenAll( _
        client.GetAsync(CStr(endpoints(0))), _
        client.GetAsync(CStr(endpoints(1))), _
        client.GetAsync(CStr(endpoints(2)))).Await
    overlappedElapsed = ElapsedSeconds(started)

    started = Timer
    For Each resource In endpoints
        Set ignored = client.GetAsync(CStr(resource)).Await
    Next resource
    sequentialElapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_REQUESTS
        .Range("C6").Value2 = overlappedElapsed
        .Range("D6").Value2 = sequentialElapsed
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
