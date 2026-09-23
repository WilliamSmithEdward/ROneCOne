VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalesRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
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
' A plain sales row used by the Excel Tables tutorial
' ----------------------------------------------------------------------------
' An everyday VBA class with nothing special in it: three properties whose
' names match the column headers of the Sales table on the sheet. That match
' is the whole trick behind mapping in both directions. ToObjects fills these
' properties by name from each table row, and DataTableFromObjects reads them
' back out by name into a new table.
'
' Property names are matched case-insensitively through VBA's own member
' dispatch, so Amount, amount, and AMOUNT all bind to the same property.
' ============================================================================

Private mAmount As Double
Private mRegion As String
Private mRep As String

Public Property Get Region() As String
    Region = mRegion
End Property

Public Property Let Region(ByVal incoming As String)
    mRegion = incoming
End Property

Public Property Get Rep() As String
    Rep = mRep
End Property

Public Property Let Rep(ByVal incoming As String)
    mRep = incoming
End Property

Public Property Get Amount() As Double
    Amount = mAmount
End Property

Public Property Let Amount(ByVal incoming As Double)
    mAmount = incoming
End Property
