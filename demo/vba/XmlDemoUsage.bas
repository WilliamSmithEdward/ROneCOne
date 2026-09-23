Attribute VB_Name = "XmlDemoUsage"
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
' ROneCOne tutorial: XML documents and typed tables
' ----------------------------------------------------------------------------
' This demo never touches the network. It parses XML with the MSXML6 engine
' Windows already ships, queries it with XPath, and lands repeated elements
' straight into a typed DataTable, then writes a table back out as XML.
'
' The parser keeps .NET's secure posture: document type definitions are
' prohibited and external references are never resolved, so a hostile
' document is refused with a typed error instead of being fetched or
' expanded. Namespaced documents work by mapping a prefix once at parse
' time. Column typing reuses the same deterministic inference the CSV layer
' uses, so XML, CSV, and JSON always agree on what a value looks like.
'
' To run it: press Alt+F8, choose RunROneCOneXmlDemo, and click Run.
' ============================================================================

Private Const BENCHMARK_ROWS As Long = 1000
Private Const BENCHMARKS_SHEET As String = "Benchmarks"
Private Const EXAMPLES_SHEET As String = "Examples"
Private Const START_SHEET As String = "Start Here"

Public Sub RunROneCOneXmlDemo()
    Dim errorDescription As String
    Dim errNumber As Long

    On Error GoTo DemoFailure
    WriteXmlExamples
    RunXmlBenchmark
    MarkDemoPassed
    Application.Calculate
    Exit Sub

DemoFailure:
    errNumber = Err.Number
    errorDescription = Err.Description
    MarkDemoFailed errNumber, errorDescription
End Sub

Private Sub WriteXmlExamples()
    Dim books As ROneCOne
    Dim catalogXml As String
    Dim catalog As ROneCOne
    Dim feed As ROneCOne
    Dim missError As Long
    Dim trace As String

    ' A small catalog exercises attributes, nesting, and entities. The same
    ' text later becomes a typed table in one call.
    catalogXml = "<catalog><book id=""1""><title>First</title>" & _
        "<price>10.5</price></book><book id=""2"">" & _
        "<title>Second &amp; Third</title></book></catalog>"
    Set catalog = ROneCOne.Xml.Parse(catalogXml)

    ' Elements in a default namespace are invisible to plain XPath; mapping
    ' a prefix once at parse time makes them queryable.
    Set feed = ROneCOne.Xml.Parse( _
        "<root xmlns=""urn:demo""><item/><item/></root>", _
        "xmlns:p='urn:demo'")

    ' Attributes and simple child elements become typed columns; the price
    ' column infers Double, and book 2's missing price loads as a null.
    Set books = ROneCOne.Xml.DeserializeTable(catalogXml, "Books", "//book")

    ' A document carrying a DOCTYPE is refused with the typed XmlError.
    trace = "unexpected success"
    On Error Resume Next
    ROneCOne.Xml.Parse "<!DOCTYPE a []><a/>"
    missError = Err.Number
    On Error GoTo 0
    If missError = ROneCOne.XmlError Then trace = "doctype refused"

    With ThisWorkbook.Worksheets(EXAMPLES_SHEET)
        .Range("E6").Value2 = catalog.Name
        .Range("E7").Value2 = CLng( _
            catalog.Elements("book").Item(0).GetAttribute("id"))
        .Range("E8").Value2 = catalog.SelectNodes("//book").Count
        .Range("E9").Value2 = _
            catalog.SelectSingleNode("//book[@id='2']/title").Value
        .Range("E10").Value2 = (catalog.SelectSingleNode("//missing") Is Nothing)
        .Range("E11").Value2 = feed.SelectNodes("//p:item").Count
        .Range("E12").Value2 = books.Rows.Item(0).Item("price")
        .Range("E13").Value2 = (InStr(1, books.ToXml(), "<Books>") > 0)
        .Range("E14").Value2 = trace
    End With
End Sub

Private Sub RunXmlBenchmark()
    Dim builder As ROneCOne
    Dim elapsed As Double
    Dim idx As Long
    Dim started As Double
    Dim parsed As ROneCOne

    ' Build a thousand-row document with the StringBuilder, then land it as
    ' a typed table in one call and count what arrived.
    Set builder = ROneCOne.StringBuilder()
    builder.Append "<rows>"
    For idx = 1 To BENCHMARK_ROWS
        builder.AppendFormat _
            "<row><id>{0}</id><total>{1:F2}</total></row>", _
            idx, idx * 1.5
    Next idx
    builder.Append "</rows>"

    started = Timer
    Set parsed = ROneCOne.Xml.DeserializeTable(builder.ToString, "Rows")
    elapsed = ElapsedSeconds(started)

    With ThisWorkbook.Worksheets(BENCHMARKS_SHEET)
        .Range("B6").Value2 = BENCHMARK_ROWS
        .Range("C6").Value2 = elapsed
        .Range("D6").Value2 = parsed.Rows.Count
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
